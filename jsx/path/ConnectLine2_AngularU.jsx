#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ConnectToAngularU.jsx

### 概要

- アンカーポイントが2つの直線パス（オープン）を複数選択し、「コの字」構造となる3本の組み合わせを検出して1つの連続パスに連結するIllustrator用スクリプトです。
- 元の3本のパスは削除され、新しい「コの字」パスが生成されます。

### 主な機能

- オープンパス3本の組み合わせから2つの交点を検出
- 水平・垂直線の自動判定とコの字構造の再構築
- 太い方のパス属性（線幅・カラー）を新パスに適用
- 元のパスを削除して整理
- 日本語／英語インターフェース対応

### 処理の流れ

1. ロック・非表示でない2点オープンパスを収集
2. 交点が2つ存在する3本のパスを検出
3. コの字型に再構築した新しいパスを生成
4. 元の3本のパスを削除

### 更新履歴

- v1.0.0 (20250613) : 初期バージョン
- v1.0.1 (20250613) : コの字パス作成処理追加、不要処理削除

---

### Script Name:

ConnectToAngularU.jsx

### Overview

- An Illustrator script that connects three selected open straight-line paths (each with two anchor points) into a single U-shaped (angular) continuous path.
- The original three paths are removed and a new U-shaped path is generated.

### Main Features

- Detects two intersections from a combination of three open paths
- Automatically identifies horizontal and vertical lines to reconstruct the U shape
- Applies thicker path attributes (stroke width and color) to the new path
- Deletes original paths for cleanup
- Japanese and English UI support

### Process Flow

1. Collect non-locked, visible two-point open paths
2. Detect three paths with two intersections
3. Create a new U-shaped path from those paths
4. Delete the original three paths

### Update History

- v1.0.0 (20250613): Initial version
- v1.0.1 (20250613): Added U-shaped path creation, removed unused processes
*/

function main() {
  try {
    if (app.documents.length === 0) {
      alert("ドキュメントが開かれていません。");
      return;
    }

    /* 線分同士の交差判定関数 */
    function segmentsIntersect(lineA, lineB) {
      var a1 = lineA.pathPoints[0].anchor;
      var a2 = lineA.pathPoints[1].anchor;
      var b1 = lineB.pathPoints[0].anchor;
      var b2 = lineB.pathPoints[1].anchor;

      function ccw(p1, p2, p3) {
        return (p3[1] - p1[1]) * (p2[0] - p1[0]) > (p2[1] - p1[1]) * (p3[0] - p1[0]);
      }

      return (ccw(a1, b1, b2) != ccw(a2, b1, b2)) && (ccw(a1, a2, b1) != ccw(a1, a2, b2));
    }

    /* 交差点の座標を取得する関数 */
    function getIntersectionPoint(lineA, lineB) {
      var a1 = lineA.pathPoints[0].anchor;
      var a2 = lineA.pathPoints[1].anchor;
      var b1 = lineB.pathPoints[0].anchor;
      var b2 = lineB.pathPoints[1].anchor;

      var A1 = a2[1] - a1[1];
      var B1 = a1[0] - a2[0];
      var C1 = A1 * a1[0] + B1 * a1[1];

      var A2 = b2[1] - b1[1];
      var B2 = b1[0] - b2[0];
      var C2 = A2 * b1[0] + B2 * b1[1];

      var det = A1 * B2 - A2 * B1;
      if (det === 0) {
        return null; // 並行または重なっている
      } else {
        var x = (B2 * C1 - B1 * C2) / det;
        var y = (A1 * C2 - A2 * C1) / det;
        return [x, y];
      }
    }

    /* 2点間の距離を測定する関数（共通化） */
    function getDistance(p1, p2) {
      var dx = p2[0] - p1[0];
      var dy = p2[1] - p1[1];
      return Math.sqrt(dx * dx + dy * dy);
    }

    var allPaths = [];
    var allPageItems = app.activeDocument.pageItems;
    for (var i = 0; i < allPageItems.length; i++) {
      var item = allPageItems[i];
      if (!(item instanceof PathItem)) continue;
      if (item.locked || item.hidden) continue;
      if (item.pathPoints.length != 2) continue;
      if (item.closed) continue;
      allPaths.push(item);
    }

    if (allPaths.length < 3) {
      alert("アンカーポイントが2つのオープンパスを3本以上配置してください。");
      return;
    }

    var lines = allPaths;
    // 「2つの交点がある3本のパスを探して、対象をその3本に限定」
    // --- 自動的に2つの交点がある3本を抽出する ---
    var triplets = [];
    for (var i = 0; i < lines.length - 2; i++) {
      for (var j = i + 1; j < lines.length - 1; j++) {
        for (var k = j + 1; k < lines.length; k++) {
          var a = lines[i];
          var b = lines[j];
          var c = lines[k];
          var count = 0;
          if (segmentsIntersect(a, b)) count++;
          if (segmentsIntersect(a, c)) count++;
          if (segmentsIntersect(b, c)) count++;
          if (count === 2) {
            triplets.push([a, b, c]);
          }
        }
      }
    }

    // すべての有効な triplet に対して順に「コの字」パスを作成
    for (var t = 0; t < triplets.length; t++) {
      var group = triplets[t];
      var horizTemp = [];
      var vertTemp = [];

      for (var i = 0; i < group.length; i++) {
        var pt1 = group[i].pathPoints[0].anchor;
        var pt2 = group[i].pathPoints[1].anchor;
        if (Math.abs(pt1[1] - pt2[1]) < 0.01) {
          horizTemp.push(group[i]);
        } else if (Math.abs(pt1[0] - pt2[0]) < 0.01 && Math.abs(pt1[1] - pt2[1]) > 10) {
          vertTemp.push(group[i]);
        }
      }

      if ((horizTemp.length >= 2 && vertTemp.length >= 1) || (horizTemp.length >= 1 && vertTemp.length >= 2)) {

        if (horizTemp.length >= 2 && vertTemp.length >= 1) {
          var topLine = null, bottomLine = null;
          var y0 = (horizTemp[0].pathPoints[0].anchor[1] + horizTemp[0].pathPoints[1].anchor[1]) / 2;
          var y1 = (horizTemp[1].pathPoints[0].anchor[1] + horizTemp[1].pathPoints[1].anchor[1]) / 2;
          topLine = y0 > y1 ? horizTemp[0] : horizTemp[1];
          bottomLine = y0 > y1 ? horizTemp[1] : horizTemp[0];
          var verticalLine = vertTemp[0];

          var ptTop = getIntersectionPoint(topLine, verticalLine);
          var ptBottom = getIntersectionPoint(bottomLine, verticalLine);

          var topP1 = topLine.pathPoints[0].anchor;
          var topP2 = topLine.pathPoints[1].anchor;
          var bottomP1 = bottomLine.pathPoints[0].anchor;
          var bottomP2 = bottomLine.pathPoints[1].anchor;

          var topDist1 = getDistance(topP1, ptTop);
          var topDist2 = getDistance(topP2, ptTop);
          var bottomDist1 = getDistance(bottomP1, ptBottom);
          var bottomDist2 = getDistance(bottomP2, ptBottom);

          if (!ptTop || !ptBottom || ptTop[1] === ptBottom[1]) {
            continue;
          }

          // 長い方の水平線を優先（長さ = 距離の2乗）
          function getLengthSq(pathItem) {
            var p1 = pathItem.pathPoints[0].anchor;
            var p2 = pathItem.pathPoints[1].anchor;
            var dx = p2[0] - p1[0];
            var dy = p2[1] - p1[1];
            return dx * dx + dy * dy;
          }

          var topLen = getLengthSq(topLine);
          var bottomLen = getLengthSq(bottomLine);
          var longerLine = topLen >= bottomLen ? topLine : bottomLine;
          var shorterLine = topLen < bottomLen ? topLine : bottomLine;

          // 左端を使う
          var longerLeft = (longerLine.pathPoints[0].anchor[0] < longerLine.pathPoints[1].anchor[0])
            ? longerLine.pathPoints[0].anchor : longerLine.pathPoints[1].anchor;
          var shorterLeft = (shorterLine.pathPoints[0].anchor[0] < shorterLine.pathPoints[1].anchor[0])
            ? shorterLine.pathPoints[0].anchor : shorterLine.pathPoints[1].anchor;

          var ptA = (longerLine === topLine) ? ptTop : ptBottom;
          var ptB = (longerLine === topLine) ? ptBottom : ptTop;

          var useTopPt = (topDist1 < topDist2) ? topP2 : topP1;
          var useBottomPt = (bottomDist1 < bottomDist2) ? bottomP2 : bottomP1;

          var newPath = app.activeDocument.pathItems.add();
          newPath.stroked = true;
          newPath.closed = false;
          newPath.setEntirePath([
            useBottomPt,
            ptBottom,
            ptTop,
            useTopPt
          ]);
          // 3本の元パスからstrokeWidth最大のものをnewPathに設定
          var all3 = [topLine, bottomLine, verticalLine];
          var refLine = all3[0];
          for (var i3 = 1; i3 < all3.length; i3++) {
            if (all3[i3].strokeWidth > refLine.strokeWidth) refLine = all3[i3];
          }
          newPath.strokeWidth = refLine.strokeWidth;
          newPath.strokeColor = refLine.strokeColor;

          // 元のパスを削除
          try { topLine.remove(); } catch (e) {}
          try { bottomLine.remove(); } catch (e) {}
          try { verticalLine.remove(); } catch (e) {}

        } else if (horizTemp.length >= 1 && vertTemp.length >= 2) {
          var leftLine = null, rightLine = null;
          var x0 = (vertTemp[0].pathPoints[0].anchor[0] + vertTemp[0].pathPoints[1].anchor[0]) / 2;
          var x1 = (vertTemp[1].pathPoints[0].anchor[0] + vertTemp[1].pathPoints[1].anchor[0]) / 2;
          leftLine = x0 < x1 ? vertTemp[0] : vertTemp[1];
          rightLine = x0 < x1 ? vertTemp[1] : vertTemp[0];
          var horizontalLine = horizTemp[0];

          var ptLeft = getIntersectionPoint(leftLine, horizontalLine);
          var ptRight = getIntersectionPoint(rightLine, horizontalLine);

          if (!ptLeft || !ptRight || ptLeft[0] === ptRight[0]) {
            continue;
          }

          var leftP1 = leftLine.pathPoints[0].anchor;
          var leftP2 = leftLine.pathPoints[1].anchor;
          var rightP1 = rightLine.pathPoints[0].anchor;
          var rightP2 = rightLine.pathPoints[1].anchor;
          var horizP1 = horizontalLine.pathPoints[0].anchor;
          var horizP2 = horizontalLine.pathPoints[1].anchor;

          var leftDist1 = getDistance(leftP1, ptLeft);
          var leftDist2 = getDistance(leftP2, ptLeft);
          var rightDist1 = getDistance(rightP1, ptRight);
          var rightDist2 = getDistance(rightP2, ptRight);

          var useLeftPt = (leftDist1 < leftDist2) ? leftP2 : leftP1;
          var useRightPt = (rightDist1 < rightDist2) ? rightP2 : rightP1;

          var newPath = app.activeDocument.pathItems.add();
          newPath.stroked = true;
          newPath.closed = false;
          newPath.setEntirePath([
            useLeftPt,
            ptLeft,
            ptRight,
            useRightPt
          ]);
          // 3本の元パスからstrokeWidth最大のものをnewPathに設定
          var all3 = [leftLine, rightLine, horizontalLine];
          var refLine = all3[0];
          for (var i3 = 1; i3 < all3.length; i3++) {
            if (all3[i3].strokeWidth > refLine.strokeWidth) refLine = all3[i3];
          }
          newPath.strokeWidth = refLine.strokeWidth;
          newPath.strokeColor = refLine.strokeColor;

          // 元のパスを削除
          try { horizontalLine.remove(); } catch (e) {}
          try { leftLine.remove(); } catch (e) {}
          try { rightLine.remove(); } catch (e) {}
        }
      }
    }

    // --- 「コの字」パス作成処理 ここまで ---

  } catch (e) {
    alert("エラーが発生しました：\n" + e);
  }
}

main();