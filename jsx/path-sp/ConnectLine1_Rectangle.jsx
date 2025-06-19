#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：ConnectToLShape.jsx

  【概要】
  交差する4本の線を検出し、それらが長方形を構成する場合は新たな長方形パスを作成します。
  元の線は削除され、生成された長方形パスが選択状態になります。
  この処理は後続の「線の連結・再構築スクリプト」と連携する前提で、対象パスの構築を補助することを目的としています。

  【処理の流れ】
  1. 水平・垂直な線を抽出
  2. 交差する4本の線の組み合わせを検出
  3. 長方形を構成可能であれば新しいパスを作成し、元の線を削除

  【対象】
  - パスアイテム（直線2点から構成される水平線・垂直線）

  【除外条件】
  - 線が4本未満の場合
  - 交差しない場合（長方形を構成できない）

  【更新履歴】
  - 2025-06-12 初版作成
  - 2025-06-15 コメント整理・drawCenterLineFromRect 処理の削除
  - 2025-06-15 JSX警告の無効化設定を追加
*/

// 4線から長方形を構成する
function attemptCreateRectangleFromLines(lines) {
  if (lines.length !== 4) return false;

  var doc = app.activeDocument;

  function getCoords(line) {
    var p1 = line.pathPoints[0].anchor;
    var p2 = line.pathPoints[1].anchor;
    return { x1: p1[0], y1: p1[1], x2: p2[0], y2: p2[1] };
  }

  var coords = [];
  for (var i = 0; i < lines.length; i++) {
    coords.push(getCoords(lines[i]));
  }

  var horizontal = [];
  var vertical = [];

  for (var i = 0; i < coords.length; i++) {
    var c = coords[i];
    if (Math.abs(c.y1 - c.y2) < 0.01) {
      horizontal.push({ index: i, y: c.y1, x1: c.x1, x2: c.x2, line: lines[i] });
    } else if (Math.abs(c.x1 - c.x2) < 0.01) {
      vertical.push({ index: i, x: c.x1, y1: c.y1, y2: c.y2, line: lines[i] });
    }
  }

  if (horizontal.length !== 2 || vertical.length !== 2) return false;

  horizontal.sort(function(a, b) { return a.y - b.y; });
  vertical.sort(function(a, b) { return a.x - b.x; });

  var topY = horizontal[1].y;
  var bottomY = horizontal[0].y;
  var leftX = vertical[0].x;
  var rightX = vertical[1].x;

  // Determine the thickest line to use as reference
  var maxStrokeWidth = 0;
  var referenceLine = lines[0];
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].strokeWidth > maxStrokeWidth) {
      maxStrokeWidth = lines[i].strokeWidth;
      referenceLine = lines[i];
    }
  }

  var newRect = doc.pathItems.add();
  newRect.stroked = true;
  newRect.filled = false;

  newRect.strokeWidth = referenceLine.strokeWidth;
  newRect.strokeColor = referenceLine.strokeColor;

  var pts = [
    [leftX, topY],
    [rightX, topY],
    [rightX, bottomY],
    [leftX, bottomY]
  ];

  for (var j = 0; j < 4; j++) {
    var pt = newRect.pathPoints.add();
    pt.anchor = pts[j];
    pt.leftDirection = pts[j];
    pt.rightDirection = pts[j];
    pt.pointType = PointType.CORNER;
  }

  newRect.closed = true;

  for (var k = 0; k < lines.length; k++) {
    lines[k].remove();
  }

  doc.selection = null;

  return true;
}

function main() {
  var doc = app.activeDocument;

  try {
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
      alert("長方形を1つ以上選択してください。");
      return;
    }

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
        return null;
      } else {
        var x = (B2 * C1 - B1 * C2) / det;
        var y = (A1 * C2 - A2 * C1) / det;
        return [x, y];
      }
    }

    var selectedItems = app.activeDocument.selection;
    var newLines = [];
    for (var i = 0; i < selectedItems.length; i++) {
      var item = selectedItems[i];
      if (
        item.typename === "PathItem" &&
        item.pathPoints.length === 2 &&
        (Math.abs(item.pathPoints[0].anchor[0] - item.pathPoints[1].anchor[0]) < 0.01 || Math.abs(item.pathPoints[0].anchor[1] - item.pathPoints[1].anchor[1]) < 0.01)
      ) {
        newLines.push(item);
      }
    }

    if (newLines.length > 0) {
      app.activeDocument.selection = newLines;
      if (newLines.length >= 4) {
        var hLines = [];
        var vLines = [];

        for (var i = 0; i < newLines.length; i++) {
          var line = newLines[i];
          var p1 = line.pathPoints[0].anchor;
          var p2 = line.pathPoints[1].anchor;
          if (Math.abs(p1[1] - p2[1]) < 0.01) {
            hLines.push(line);
          } else if (Math.abs(p1[0] - p2[0]) < 0.01) {
            vLines.push(line);
          }
        }

        for (var i = 0; i < hLines.length - 1; i++) {
          for (var j = i + 1; j < hLines.length; j++) {
            var hy1 = hLines[i];
            var hy2 = hLines[j];
            for (var m = 0; m < vLines.length - 1; m++) {
              for (var n = m + 1; n < vLines.length; n++) {
                var vx1 = vLines[m];
                var vx2 = vLines[n];

                if (
                  segmentsIntersect(hy1, vx1) &&
                  segmentsIntersect(hy1, vx2) &&
                  segmentsIntersect(hy2, vx1) &&
                  segmentsIntersect(hy2, vx2)
                ) {
                  var candidateLines = [hy1, hy2, vx1, vx2];
                  var created = attemptCreateRectangleFromLines(candidateLines);
                  if (created) return;
                }
              }
            }
          }
        }
      }
    }
  } catch (e) {
    alert("エラーが発生しました：\n" + e);
  }
}

main();