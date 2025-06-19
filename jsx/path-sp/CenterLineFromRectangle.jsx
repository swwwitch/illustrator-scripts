#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名: CenterLineFromRect.jsx

  概要:
  選択した長方形の中心に縦線または横線を描画します。
  描画された線は、同一フォルダ内のスクリプトに引き継がれて連結・再構築処理されます。

  処理の流れ:
  1. 選択オブジェクトの中から、閉じたパスでアンカーポイントが4点の長方形を抽出
  2. 対象長方形の縦横比に基づいて、中心線（縦または横）を描画
  3. 元の長方形を削除し、描画された線のみを選択状態にする

  対象:
  - PathItem（閉じたパスでアンカーポイントが4点の長方形）
  - CompoundPathItem 内の同様の長方形

  除外条件:
  - 正方形に近い形状（縦横差が5%未満）
  - 短辺 × 1.5 ＞ 長辺 の形状

  更新履歴:
  - 2025-06-12 初版作成（縦長・横長対応、複数オブジェクト対応、CMYK割れ判定付き）
  - 2025-06-15 除外条件追加・コメント整理
  - 2025-06-15 中心線描画機能のみを残し、長方形再構築ロジックを削除
  - 2025-06-15 回転補正ロジックをアンカーポイントベースに変更、短辺 × 1.8 を 1.5 に修正
*/

// 長方形の中心に線を描画し、元の長方形を削除する関数
function drawCenterLineFromRect(rect) {
  // --- 回転補正（最初の2点のアンカーポイントから角度を取得） ---
  if (rect.pathPoints.length === 4) {
    var ptA = rect.pathPoints[0].anchor;
    var ptB = rect.pathPoints[1].anchor;

    var dx = ptB[0] - ptA[0];
    var dy = ptB[1] - ptA[1];
    var angleRad = Math.atan2(dy, dx);
    var angleDeg = angleRad * 180 / Math.PI;
    if (angleDeg < 0) angleDeg += 360;

    var rotationAmount = angleDeg;
    var normalized = rotationAmount % 90;
    if (normalized > 45) normalized = 90 - normalized;

    if (normalized >= 0.5 && normalized <= 10) {
      rect.rotate(-rotationAmount);
    }
  }

  // --- 補正後にバウンディングボックスを取得 ---
  var b = rect.geometricBounds;
  var left = b[0], top = b[1], right = b[2], bottom = b[3];
  var w = right - left;
  var h = top - bottom;

  // 正方形に近い場合は除外（5%未満の差）
  var diffRatio = Math.abs(w - h) / Math.max(w, h);
  if (diffRatio < 0.05) return null;
  // 短辺 × 1.5 ＞ 長辺の場合は除外
  var shortSide = Math.min(w, h);
  var longSide = Math.max(w, h);
  if (shortSide * 1.5 > longSide) return null;

  var centerLine = app.activeDocument.pathItems.add();
  centerLine.stroked = true;
  centerLine.filled = false;
  centerLine.strokeColor = rect.fillColor;

  var p1 = centerLine.pathPoints.add();
  var p2 = centerLine.pathPoints.add();

  if (h <= w) {
    // 横長：中央に水平線を描く
    var centerY = (top + bottom) / 2;
    p1.anchor = [left, centerY];
    p2.anchor = [right, centerY];
  } else {
    // 縦長：中央に垂直線を描く
    var centerX = (left + right) / 2;
    p1.anchor = [centerX, top];
    p2.anchor = [centerX, bottom];
  }

  centerLine.strokeWidth = (h <= w) ? h : w;

  p1.leftDirection = p1.anchor;
  p1.rightDirection = p1.anchor;
  p2.leftDirection = p2.anchor;
  p2.rightDirection = p2.anchor;

  rect.remove();
  return centerLine;
}

function main() {
  var doc = app.activeDocument;

  try {
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
      alert("長方形を1つ以上選択してください。");
      return;
    }

    var selectedItems = app.activeDocument.selection;
    var newLines = [];

    for (var i = 0; i < selectedItems.length; i++) {
      var target = selectedItems[i];

      if (target.typename === "PathItem" && target.closed && target.pathPoints.length === 4) {
        var newLine = drawCenterLineFromRect(target);
        if (newLine) newLines.push(newLine);
      } else if (target.typename === "CompoundPathItem" && target.pathItems.length === 1) {
        var subPath = target.pathItems[0];
        if (subPath.closed && subPath.pathPoints.length === 4) {
          var newLine = drawCenterLineFromRect(subPath);
          if (newLine) newLines.push(newLine);
        }
      }
    }
    if (newLines.length > 0) {
      app.activeDocument.selection = newLines;
    }
  } catch (e) {
    alert("エラーが発生しました：\n" + e);
  }
}

main();