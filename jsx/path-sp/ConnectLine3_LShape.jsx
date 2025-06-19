#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  ConnectToLShape.jsx

  概要：
  ドキュメント内のロック・非表示でないオープンパス同士で交差するものを見つけ、
  交点を起点に「コの字（L字型）」に連結するスクリプトです。元の2本のパスは削除されます。

  処理の流れ：
  1. ドキュメント内のロック・非表示でない PathItem を収集
  2. オープンパスのみ抽出
  3. 交差点を持つ2本のオープンパスの交点を検出
  4. 各パスに交点を挿入し、交点を起点にL字型パスを作成
  5. 元パスを削除し、新しいL字型パスを描画

  更新日：2025-06-15
*/

main();

function main() {
    var processed = false;
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    // ドキュメント内のロック・非表示でない PathItem を収集
    var allPaths = [];
    var allPageItems = app.activeDocument.pageItems;

    for (var i = 0; i < allPageItems.length; i++) {
        var item = allPageItems[i];
        if (
            item.typename === "PathItem" &&
            !item.locked &&
            !item.hidden &&
            !item.layer.locked &&
            !item.layer.hidden
        ) {
            allPaths.push(item);
        }
    }

    // オープンパスのみ抽出
    var openPaths = [];
    for (var i = 0; i < allPaths.length; i++) {
        if (isOpenPath(allPaths[i])) openPaths.push(allPaths[i]);
    }

    if (openPaths.length < 2) {
        alert("有効なオープンパスが2本以上必要です。");
        return;
    }

    // 交差して連結したペアは対象から除外しながら進める
    while (openPaths.length >= 2) {
        var foundPair = false;

        for (var i = 0; i < openPaths.length - 1; i++) {
            for (var j = i + 1; j < openPaths.length; j++) {
                var pathA = openPaths[i];
                var pathB = openPaths[j];
                var intersections = getIntersections(pathA, pathB);

                if (intersections.length > 0) {
                    for (var k = 0; k < intersections.length; k++) {
                        insertAnchorPointAt(pathA, intersections[k]);
                        insertAnchorPointAt(pathB, intersections[k]);
                    }

                    createLShapedPath(pathA, pathB, intersections[0]);
                    try { pathA.remove(); } catch (e) {}
                    try { pathB.remove(); } catch (e) {}

                    // 配列から除外
                    openPaths.splice(j, 1);
                    openPaths.splice(i, 1);
                    foundPair = true;
                    processed = true;
                    break;
                }
            }
            if (foundPair) break;
        }

        if (!foundPair) break;
    }

    if (!processed) {
        alert("交差している2本のパスが見つかりませんでした。");
    }
}

// 長いパスを主軸に交点でL字型パスを作成し元パスを削除
function createLShapedPath(pathA, pathB, intersection) {
    try {
        if (!pathA.pathPoints || !pathB.pathPoints) return;
    } catch (e) {
        return;
    }

    var lengthA = pathLength(pathA);
    var lengthB = pathLength(pathB);
    var mainPath = lengthA >= lengthB ? pathA : pathB;
    var subPath = lengthA >= lengthB ? pathB : pathA;

    var ptsMain = extractSegmentAwayFromIntersection(mainPath, intersection);
    var ptsSub = extractSegmentAwayFromIntersection(subPath, intersection);

    if (!pointsEqual(ptsMain[ptsMain.length - 1], intersection)) {
        ptsMain.push(intersection);
    }
    if (!pointsEqual(ptsSub[0], intersection)) {
        ptsSub.unshift(intersection);
    }

    var newPathPoints = ptsMain.concat(ptsSub);

    // どちらのパスの線幅が太いかを比較
    var strokeWidthA = pathA.strokeWidth;
    var strokeWidthB = pathB.strokeWidth;
    var refPath = (strokeWidthA >= strokeWidthB) ? pathA : pathB;

    // 新しいパスを作成
    var newPath = app.activeDocument.pathItems.add();
    newPath.setEntirePath(newPathPoints);
    newPath.stroked = true;
    newPath.filled = false;

    // 太い方のパスの線幅とカラーを適用
    newPath.strokeWidth = refPath.strokeWidth;
    newPath.strokeColor = refPath.strokeColor;

}

// 交点から遠ざかる方向のセグメントを抽出
function extractSegmentAwayFromIntersection(path, intersection) {
    var pts = path.pathPoints;
    var startPt = pts[0].anchor;
    var endPt = pts[pts.length - 1].anchor;
    var distStart = distance(startPt, intersection);
    var distEnd = distance(endPt, intersection);
    var segment = [];

    if (distStart >= distEnd) {
        for (var i = 0; i < pts.length; i++) {
            segment.push(pts[i].anchor);
            if (distance(pts[i].anchor, intersection) < 0.5) break;
        }
    } else {
        for (var i = pts.length - 1; i >= 0; i--) {
            segment.push(pts[i].anchor);
            if (distance(pts[i].anchor, intersection) < 0.5) break;
        }
        segment.reverse();
    }
    return segment;
}

// 座標がほぼ等しいか判定
function pointsEqual(p1, p2) {
    return (Math.abs(p1[0] - p2[0]) < 0.5 && Math.abs(p1[1] - p2[1]) < 0.5);
}

// パスが開いているかどうかを判定
function isOpenPath(path) {
    return path.typename === "PathItem" && !path.closed;
}

// 2つのパスの交差点を取得
function getIntersections(pathA, pathB) {
    var points = [];
    var segsA = getSegments(pathA);
    var segsB = getSegments(pathB);

    for (var i = 0; i < segsA.length; i++) {
        for (var j = 0; j < segsB.length; j++) {
            var pt = getIntersection(segsA[i], segsB[j]);
            if (pt) points.push(pt);
        }
    }
    return points;
}

// パスのセグメント配列を取得
function getSegments(path) {
    var pts = path.pathPoints;
    var segs = [];
    for (var i = 0; i < pts.length - 1; i++) {
        segs.push({
            p1: pts[i].anchor,
            p2: pts[i + 1].anchor
        });
    }
    return segs;
}

// 2つのセグメントの交点を計算
function getIntersection(seg1, seg2) {
    var x1 = seg1.p1[0], y1 = seg1.p1[1];
    var x2 = seg1.p2[0], y2 = seg1.p2[1];
    var x3 = seg2.p1[0], y3 = seg2.p1[1];
    var x4 = seg2.p2[0], y4 = seg2.p2[1];

    var denom = (x1 - x2)*(y3 - y4) - (y1 - y2)*(x3 - x4);
    if (denom === 0) return null;

    var px = ((x1*y2 - y1*x2)*(x3 - x4) - (x1 - x2)*(x3*y4 - y3*x4)) / denom;
    var py = ((x1*y2 - y1*x2)*(y3 - y4) - (y1 - y2)*(x3*y4 - y3*x4)) / denom;

    if (isBetween(px, x1, x2) && isBetween(py, y1, y2) &&
        isBetween(px, x3, x4) && isBetween(py, y3, y4)) {
        return [px, py];
    }
    return null;
}

// 値が区間内にあるか判定
function isBetween(val, a, b) {
    return (val >= Math.min(a, b) - 0.1) && (val <= Math.max(a, b) + 0.1);
}

// 指定点にアンカーポイントを挿入
function insertAnchorPointAt(path, point) {
    var pts;
    try {
        pts = path.pathPoints;
    } catch (e) {
        return; // 無効な path の場合は処理しない
    }

    var idx = findClosestSegmentIndex(path, point);
    if (idx === -1) return;

    var newPt = pts.add();
    for (var i = pts.length - 1; i > idx; i--) {
        pts[i].anchor = pts[i - 1].anchor;
        pts[i].leftDirection = pts[i - 1].leftDirection;
        pts[i].rightDirection = pts[i - 1].rightDirection;
        pts[i].pointType = pts[i - 1].pointType;
    }
    pts[idx + 1].anchor = point;
    pts[idx + 1].leftDirection = point;
    pts[idx + 1].rightDirection = point;
    pts[idx + 1].pointType = PointType.CORNER;
}

// 点がセグメント上にあるか判定し、該当セグメントのインデックスを返す
function findClosestSegmentIndex(path, point) {
    var pts = path.pathPoints;
    for (var i = 0; i < pts.length - 1; i++) {
        var seg = {p1: pts[i].anchor, p2: pts[i + 1].anchor};
        if (isPointOnSegment(point, seg)) {
            return i;
        }
    }
    return -1;
}

// 点がセグメント上にあるか判定
function isPointOnSegment(pt, seg) {
    var d1 = distance(pt, seg.p1);
    var d2 = distance(pt, seg.p2);
    var d = distance(seg.p1, seg.p2);
    return Math.abs((d1 + d2) - d) < 0.5;
}

// 2点間の距離を計算
function distance(p1, p2) {
    return Math.sqrt(Math.pow(p2[0] - p1[0], 2) + Math.pow(p2[1] - p1[1], 2));
}

// パスの長さを計算
function pathLength(path) {
    var pts = path.pathPoints;
    var len = 0;
    for (var i = 0; i < pts.length - 1; i++) {
        len += distance(pts[i].anchor, pts[i + 1].anchor);
    }
    return len;
}