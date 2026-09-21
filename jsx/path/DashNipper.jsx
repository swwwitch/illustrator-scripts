#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

破線を選択して実行すると、アウトライン化して1つずつの線分に分け、それぞれを同じ見た目の線（中心線）に置き換えて、破線ごとにグループにまとめます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nae6882ac8a73

### Overview

With a dashed line selected, outlines it, splits it into its dashes, replaces each dash with a line that looks the same (its center line), and groups the lines of each dashed line.

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

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashNipper.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nae6882ac8a73"; /* 紹介記事 / article URL */

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
            notDashedLine: { ja: "破線を選択してください。", en: "Select a dashed line." },
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
    var TINY_SEGMENT_RATIO = 0.05;        /* 線幅に対して、ごく短いとみなすセグメントの比率 / segment length treated as tiny (ratio to stroke width) */

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
     * 2つのアンカーを結ぶセグメントが直線かどうかを返す（ハンドルが両端を結ぶ線上にあれば直線）
     * @param {AnchorInfo} from - セグメントの始点
     * @param {AnchorInfo} to - セグメントの終点
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(from, to) {
        return getDistanceToLine(from.rightDirection, from.anchor, to.anchor) <= STRAIGHT_TOLERANCE &&
            getDistanceToLine(to.leftDirection, from.anchor, to.anchor) <= STRAIGHT_TOLERANCE;
    }

    /**
     * パスの1辺が直線かどうかを返す
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} sideIndex - 辺の始点になるアンカーの番号
     * @returns {boolean} 直線なら true
     */
    function isStraightSide(anchors, sideIndex) {
        return isStraightSegment(anchors[sideIndex % anchors.length], anchors[(sideIndex + 1) % anchors.length]);
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
     * @param {number} expectedThickness - 元の破線の線幅（pt）。半円の直径と一致するものだけを使う
     * @returns {CenterLine|null} 線端を丸型にする中心線。ピル形状でなければ null
     */
    function buildPillCenterLine(anchors, expectedThickness) {
        var anchorCount = anchors.length;
        for (var capStart = 0; capStart < anchorCount; capStart++) {
            for (var capSegments = 1; capSegments <= MAX_CAP_SEGMENTS; capSegments++) {
                if (!isSemicircleChain(anchors, capStart, capSegments)) continue;
                var capDiameter = getDistance(anchors[capStart].anchor, anchors[(capStart + capSegments) % anchorCount].anchor);
                if (!isMatchingLength(capDiameter, expectedThickness)) continue;

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
     * そうでなければ、太さ（両端の辺の長さ）が線幅と一致する向きを使う（線分は線幅より短いことがあり、
     * 細い方を選ぶと向きを取り違えるため）。アンカーが4つのときはどの形でも受け付け、
     * それより多いときは両端が直線で長さのそろうものだけを受け付ける
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} expectedThickness - 元の破線の線幅（pt）
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
            if (!isMatchingLength(candidate.thickness, expectedThickness)) continue;
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
    // アウトラインの整形 / Outline cleanup
    // =========================================
    // 閉じたパスの破線をアウトライン化すると、継ぎ目の近くの線分に、長さ 0.1pt 前後のセグメントや
    // 切り口の途中のアンカーが混ざる。中心線を求められなかった線分だけ、これを取り除いてから判定し直す
    // Outlining a closed dashed path leaves ~0.1pt segments and extra anchors on the cut near the seam.
    // Only dashes that fail the first attempt are cleaned up and tried again

    /**
     * 近接したアンカーの集まりを1つのアンカーにまとめる（位置は平均、ハンドルは位置の移動に合わせてずらす）
     * @param {AnchorInfo[]} cluster - まとめるアンカー（パスの順）
     * @returns {AnchorInfo} まとめたアンカー
     */
    function mergeAnchorCluster(cluster) {
        if (cluster.length === 1) return cluster[0];
        var sumX = 0;
        var sumY = 0;
        for (var i = 0; i < cluster.length; i++) {
            sumX += cluster[i].anchor[0];
            sumY += cluster[i].anchor[1];
        }
        var position = [sumX / cluster.length, sumY / cluster.length];
        var first = cluster[0];
        var last = cluster[cluster.length - 1];
        return {
            anchor: position,
            leftDirection: [first.leftDirection[0] + position[0] - first.anchor[0], first.leftDirection[1] + position[1] - first.anchor[1]],
            rightDirection: [last.rightDirection[0] + position[0] - last.anchor[0], last.rightDirection[1] + position[1] - last.anchor[1]]
        };
    }

    /**
     * ごく短いセグメントでつながったアンカーを、1つのアンカーにまとめる
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} mergeDistance - これより短いセグメントをまとめる（pt）
     * @returns {AnchorInfo[]} まとめたあとのアンカー。全体がごく小さいときは空の配列
     */
    function mergeTinySegments(anchors, mergeDistance) {
        var anchorCount = anchors.length;
        /* 手前のセグメントが長いアンカーから数え始め、集まりが先頭と末尾に割れないようにする
           Start right after a long segment so no cluster wraps around the start */
        var startIndex = -1;
        for (var i = 0; i < anchorCount; i++) {
            var previousAnchor = anchors[(i - 1 + anchorCount) % anchorCount];
            if (getDistance(previousAnchor.anchor, anchors[i].anchor) >= mergeDistance) {
                startIndex = i;
                break;
            }
        }
        if (startIndex < 0) return [];

        var mergedAnchors = [];
        var offset = 0;
        while (offset < anchorCount) {
            var cluster = [anchors[(startIndex + offset) % anchorCount]];
            offset++;
            while (offset < anchorCount &&
                getDistance(cluster[cluster.length - 1].anchor, anchors[(startIndex + offset) % anchorCount].anchor) < mergeDistance) {
                cluster.push(anchors[(startIndex + offset) % anchorCount]);
                offset++;
            }
            mergedAnchors.push(mergeAnchorCluster(cluster));
        }
        return mergedAnchors;
    }

    /**
     * 直線どうしのつなぎ目で、ほぼ一直線上にある余分なアンカーを取り除く
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} maxDeviation - 前後のアンカーを結ぶ線からのずれがこれ以下なら取り除く（pt）
     * @returns {AnchorInfo[]} 取り除いたあとのアンカー
     */
    function removeStraightMidAnchors(anchors, maxDeviation) {
        var remainingAnchors = anchors.slice(0);
        var hasRemoved = true;
        while (hasRemoved && remainingAnchors.length > 3) {
            hasRemoved = false;
            for (var i = 0; i < remainingAnchors.length; i++) {
                var anchorCount = remainingAnchors.length;
                var previousAnchor = remainingAnchors[(i - 1 + anchorCount) % anchorCount];
                var currentAnchor = remainingAnchors[i];
                var nextAnchor = remainingAnchors[(i + 1) % anchorCount];
                if (!isStraightSegment(previousAnchor, currentAnchor) || !isStraightSegment(currentAnchor, nextAnchor)) continue;
                if (getDistanceToLine(currentAnchor.anchor, previousAnchor.anchor, nextAnchor.anchor) > maxDeviation) continue;
                /* 前後のアンカーの間にあるとき（折り返しでないとき）だけ / Only when it lies between its neighbors */
                if (getDistance(previousAnchor.anchor, currentAnchor.anchor) >= getDistance(previousAnchor.anchor, nextAnchor.anchor) ||
                    getDistance(nextAnchor.anchor, currentAnchor.anchor) >= getDistance(previousAnchor.anchor, nextAnchor.anchor)) continue;
                remainingAnchors.splice(i, 1);
                hasRemoved = true;
                break;
            }
        }
        return remainingAnchors;
    }

    /**
     * 破線の線分から中心線を求める。そのままで求められなければ、ごく短いセグメントと余分なアンカーを除いて判定し直す
     * @param {AnchorInfo[]} anchors - 線分のアンカー情報
     * @param {number} strokeWidth - 元の破線の線幅（pt）
     * @returns {CenterLine|null} 中心線。求められなければ null
     */
    function getDashCenterLine(anchors, strokeWidth) {
        var centerLine = getCenterLine(anchors, strokeWidth);
        if (centerLine) return centerLine;
        var tinyLength = strokeWidth * TINY_SEGMENT_RATIO;
        return getCenterLine(removeStraightMidAnchors(mergeTinySegments(anchors, tinyLength), tinyLength), strokeWidth);
    }

    /**
     * アウトライン化で生じた、線分とは呼べないほど小さな破片かどうかを返す
     * @param {AnchorInfo[]} anchors - 線分のアンカー情報
     * @param {number} strokeWidth - 元の破線の線幅（pt）
     * @returns {boolean} ごく短いセグメントをまとめるとアンカーが3つ未満になるなら true
     */
    function isOutlineDebris(anchors, strokeWidth) {
        return mergeTinySegments(anchors, strokeWidth * TINY_SEGMENT_RATIO).length < 3;
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
        /* グループの中もたどる。文字の選択中は selection が TextRange になり、[i] は undefined なので飛ばされる
           Groups are searched too; a TextRange selection yields undefined items, which are skipped */
        var paths = collectItemsByType(selectedItems, "PathItem");
        for (var i = 0; i < paths.length; i++) {
            if (paths[i].stroked && paths[i].strokeDashes.length > 0) dashedPaths.push(paths[i]);
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
     * 元の図形を中心線に置き換える
     * @param {PathItem} sourcePath - 元の図形
     * @param {CenterLine} centerLine - 中心線
     * @returns {PathItem} 描いた中心線
     */
    function replaceWithCenterLine(sourcePath, centerLine) {
        var linePath = drawCenterLine(sourcePath, centerLine);
        sourcePath.remove();
        return linePath;
    }

    /**
     * アイテムを1つのグループにまとめる（グループは先頭のアイテムの位置に置き、アイテムの並び順は保つ）
     * @param {PageItem[]} items - まとめるアイテム（前面から順）
     * @returns {GroupItem} 作ったグループ
     */
    function groupInPlace(items) {
        var wrapperGroup = items[0].parent.groupItems.add();
        wrapperGroup.move(items[0], ElementPlacement.PLACEBEFORE);
        for (var i = 0; i < items.length; i++) {
            items[i].move(wrapperGroup, ElementPlacement.PLACEATEND);
        }
        return wrapperGroup;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 1本の破線を線分ごとの中心線に置き換え、2つ以上になればグループにまとめる
     * 中心線にできなかった線分もアウトラインのまま同じグループに入れ、アウトライン化で生じた破片は削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} dashedPath - 破線のパス
     * @returns {{resultItem: PageItem|null, skippedCount: number}} 置き換えた結果（グループか1本の線。線分が無ければ null）と、中心線にできなかった線分の数
     */
    function convertDashedPath(doc, dashedPath) {
        var strokeWidth = dashedPath.strokeWidth;
        var dashPieces = outlineDashedPath(doc, dashedPath);
        var resultItems = [];
        var skippedCount = 0;
        for (var i = 0; i < dashPieces.length; i++) {
            var anchors = readAnchors(dashPieces[i]);
            if (isOutlineDebris(anchors, strokeWidth)) {
                dashPieces[i].remove();
                continue;
            }
            var centerLine = dashPieces[i].closed ? getDashCenterLine(anchors, strokeWidth) : null;
            if (centerLine) {
                resultItems.push(replaceWithCenterLine(dashPieces[i], centerLine));
            } else {
                resultItems.push(dashPieces[i]);
                skippedCount++;
            }
        }

        var resultItem = null;
        if (resultItems.length === 1) resultItem = resultItems[0];
        if (resultItems.length > 1) resultItem = groupInPlace(resultItems);
        return { resultItem: resultItem, skippedCount: skippedCount };
    }

    /**
     * 破線を線分ごとの中心線に置き換える（破線ごとに1つのグループ）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} dashedPaths - 破線のパス
     * @returns {void}
     */
    function convertDashedPaths(doc, dashedPaths) {
        var resultItems = [];
        var skippedCount = 0;
        var hasOutlineFailure = false;
        /* 線幅は破線ごとに違いうるので、1本ずつアウトライン化する
           Outline one path at a time, since each may have its own stroke width */
        for (var i = 0; i < dashedPaths.length; i++) {
            var conversion = convertDashedPath(doc, dashedPaths[i]);
            if (conversion.resultItem) {
                resultItems.push(conversion.resultItem);
            } else {
                hasOutlineFailure = true;
            }
            skippedCount += conversion.skippedCount;
        }

        doc.selection = resultItems;
        if (hasOutlineFailure) {
            alert(getLabel(LABELS.alert.dashOutlineFailed));
        }
        if (skippedCount > 0) {
            alert(getLabel(LABELS.alert.skippedPieces).replace("{count}", skippedCount));
        }
    }

    /**
     * 選択中の破線を、1本ずつ順に線分ごとの中心線に置き換える（破線が無ければ警告を出して終了）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var dashedPaths = doc.selection ? collectDashedPaths(doc.selection) : [];
        if (dashedPaths.length === 0) {
            alert(getLabel(LABELS.alert.notDashedLine));
            return;
        }
        convertDashedPaths(doc, dashedPaths);
    }

    main();

})();
