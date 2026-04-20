function appendTopEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, rowCount, oneThirdWide, oneQuarterHigh, pieceEdgeData) {
  /* 上辺突起 */
  if (y < rowCount - 1) {
    if (pieceEdgeData[y + 1][x].topOut) {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;
    } else {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;
    }
  }
}

function appendRightEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, columnCount, oneThirdHigh, oneQuarterWide, pieceEdgeData) {
  /* 右辺突起 */
  if (x < columnCount - 1) {
    if (pieceEdgeData[y][x + 1].rightOut) {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 5.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 5.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;
    } else {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 3 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 2.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 3 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 2.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;
    }
  }
}

function appendBottomEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdWide, oneQuarterHigh, pieceEdgeData) {
  /* 下辺突起 */
  if (y > 0) {
    if (pieceEdgeData[y][x].topOut) {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;
    } else {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset];
      curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset];
      curvept.pointType = PointType.SMOOTH;
    }
  }
}

function appendLeftEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdHigh, oneQuarterWide, pieceEdgeData) {
  /* 左辺突起 */
  if (x > 0) {
    if (pieceEdgeData[y][x].rightOut) {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth + oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;
    } else {
      var curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth - 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth - 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;

      curvept = maskPath.pathPoints.add();
      curvept.anchor = [originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh];
      curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh];
      curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh];
      curvept.pointType = PointType.SMOOTH;
    }
  }
}
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartSliceWithPuzzlify.jsx

### 概要

- 選択した画像や図形を、グリッドまたはジグソーパズル形状に分割し、それぞれにマスクを適用するIllustrator用スクリプトです。
- 行数・列数・ピース数の指定に対応し、オフセット、オーバーラップ、バラけ処理、ケイ、角丸などを組み合わせて調整できます。
- 日本語／英語UIに対応し、入力欄では ↑↓ / Shift+↑↓ / Option+↑↓ による値の増減も行えます。

### 主な機能

- グリッド分割、トラディショナル、ランダムなジグソー形状の生成
- 行数・列数・ピース数の指定と、選択オブジェクト比率に応じた自動行列算出
- オフセット、オーバーラップ、散乱（バラけ）処理の適用
- ケイの追加、角丸の適用
- 画像、シンボル、ベクターアートワーク、複数選択オブジェクトへの対応
- 複数選択時の一時グループ化とシンボル化処理
- 日本語／英語インターフェース対応
- 入力欄でのキーボード増減操作
  - ↑↓キー：±1
  - Shift+↑↓：±10
  - Option+↑↓：±0.1

### 処理の流れ

1. 対象オブジェクト（画像、シンボル、ベクターなど）を選択
2. ダイアログで分割モード、ピース数、行数、列数、形状、各種効果を指定
3. モードに応じてグリッドまたはジグソー形状の各ピースを生成
4. 必要に応じてオフセットやバラけ処理を適用
5. 元オブジェクトを削除し、新しいピースを配置

### オリジナル、謝辞

Originally created by Jongware on 7-Oct-2010
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### 更新履歴

- v1.0 (20250607) : 初期バージョン
- v1.1 (20250608) : シンボル対応、ベクターアートワーク対応
- v1.2 (20250609) : グリッド形状対応
- v1.4.0 (20250610) : オフセット機能追加、単位コード対応、ダイアログUI整理、入力欄のキー操作対応

---

### Script Name:

SmartSliceWithPuzzlify.jsx

### Overview

- An Illustrator script that splits a selected image or shape into either a grid or jigsaw puzzle pieces and applies a mask to each piece.
- Supports rows, columns, and total piece count, along with offset, overlap, scatter, stroke, and round-corner adjustments.
- Supports Japanese and English UI, and numeric fields can be adjusted with ↑↓ / Shift+↑↓ / Option+↑↓ keyboard shortcuts.

### Main Features

- Generate grid, traditional, or random jigsaw-shaped pieces
- Specify rows, columns, and total pieces, with automatic row/column calculation based on the selected artwork ratio
- Apply offset, overlap, and scatter effects
- Add stroke and apply round corners
- Supports images, symbols, vector artwork, and multiple selected objects
- Temporary grouping and symbolization for multi-selection workflows
- Japanese and English UI support
- Keyboard increment/decrement support in numeric fields
  - Up/Down: ±1
  - Shift+Up/Down: ±10
  - Option+Up/Down: ±0.1

### Process Flow

1. Select a target object (image, symbol, vector artwork, etc.)
2. Configure split mode, total pieces, rows, columns, shape type, and optional effects in the dialog
3. Generate grid or jigsaw-shaped pieces according to the selected mode
4. Optionally apply offset and scatter effects
5. Remove the original object and place the generated pieces

### Original / Acknowledgements

Originally created by Jongware on 7-Oct-2010
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### Update History

- v1.0 (20250607): Initial version
- v1.1 (20250608): Added symbol and vector artwork support
- v1.2 (20250609): Added grid shape support
- v1.4.0 (20250610): Added offset support, unit-code handling, dialog UI cleanup, and keyboard increment support for numeric fields
*/

(function () {

  var SCRIPT_VERSION = "v1.4.0";

  function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
  }
  var lang = getCurrentLang();

  /* 日英ラベル定義 / Japanese-English label definitions */
  var LABELS = {
    modePuzzle: { ja: "パズル", en: "Puzzle" },
    modeGridSplit: { ja: "グリッド", en: "Grid" },
    shapePanel: { ja: "形状", en: "Shape" },
    shapeTraditional: { ja: "トラディショナル", en: "Traditional" },
    shapeRandom: { ja: "ランダム", en: "Random" },
    totalPieces: { ja: "ピース数", en: "Total pieces" },
    columns: { ja: "列数", en: "Columns" },
    rows: { ja: "行数", en: "Rows" },
    explode: { ja: "バラけ処理", en: "Explode" },
    offsetLabel: { ja: "オフセット", en: "Offset" },
    overlap: { ja: "オーバーラップ", en: "Overlap" },
    ruleCheck: { ja: "ケイ", en: "Add Stroke" },
    roundCheck: { ja: "角丸", en: "Apply Round Corners" },
    dialogTitle: {
      ja: "リンク画像の分割",
      en: "Split Linked Image"
    },
    okBtn: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
  };

  function L(key) {
    return LABELS[key][lang];
  }

  // 単位コードとラベルのマップ
  var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
  };

  // 現在の単位ラベルを取得
  function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
  }

  // 現在の単位→pt 変換係数 / Unit-to-pt conversion factor
  function getUnitToPtFactor() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    switch (unitCode) {
      case 0: return 72;               /* in */
      case 1: return 72 / 25.4;        /* mm */
      case 2: return 1;                /* pt */
      case 3: return 12;               /* pica */
      case 4: return 72 / 2.54;        /* cm */
      case 5: return 72 / 25.4 * 0.25; /* Q */
      case 6: return 1;                /* px */
      case 7: return 72;               /* ft/in */
      case 8: return 72 / 0.0254;      /* m */
      case 9: return 72 * 36;          /* yd */
      case 10: return 72 * 12;         /* ft */
      default: return 1;
    }
  }

  // 画像サイズに基づく初期グリッドサイズを計算
  function getInitialGridSize(imageWidth, imageHeight, totalPieces) {
    var aspectRatio = imageWidth / imageHeight;
    var cols = Math.round(Math.sqrt(totalPieces * aspectRatio));
    if (cols < 1) cols = 1;
    var rows = Math.round(totalPieces / cols);
    if (rows < 1) rows = 1;
    return [rows, cols];
  }

  // EditText arrow key increment/decrement support
  function changeValueByArrowKey(editText, allowNegative) {
    editText.addEventListener("keydown", function (event) {
      var value = Number(editText.text);
      if (isNaN(value)) return;

      var keyboard = ScriptUI.environment.keyboardState;
      var delta = 1;
      var handled = false;

      if (keyboard.shiftKey) {
        delta = 10;
        if (event.keyName == "Up") {
          value = Math.ceil((value + 1) / delta) * delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value = Math.floor((value - 1) / delta) * delta;
          handled = true;
        }
      } else if (keyboard.altKey) {
        delta = 0.1;
        if (event.keyName == "Up") {
          value += delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value -= delta;
          handled = true;
        }
      } else {
        delta = 1;
        if (event.keyName == "Up") {
          value += delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value -= delta;
          handled = true;
        }
      }

      if (!handled) return;

      if (keyboard.altKey) {
        value = Math.round(value * 10) / 10;
      } else {
        value = Math.round(value);
      }

      if (!allowNegative && value < 0) value = 0;

      event.preventDefault();
      editText.text = value;
    });
  }

  /* ダイアログボックス作成 / Build the dialog window */
  function createDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignment = 'right';

    /* 共通セクション（上部） / Common section (top) */
    var commonGroup = dlg.add('group');
    commonGroup.orientation = 'column';
    commonGroup.alignChildren = 'left';
    commonGroup.margins = [10, 5, 10, 5];

    /* モード選択（パズル / グリッド分割） */
    var modeGroup = commonGroup.add("group");
    modeGroup.orientation = "row";
    modeGroup.alignChildren = "left";
    var modeRadioGrid = modeGroup.add("radiobutton", undefined, LABELS.modeGridSplit[lang]);
    var modeRadioPuzzle = modeGroup.add("radiobutton", undefined, LABELS.modePuzzle[lang]);
    modeRadioGrid.value = true;

    /* 列数・行数 */
    var rowColGroup = commonGroup.add('group');
    rowColGroup.orientation = 'row';
    rowColGroup.alignChildren = 'left';
    var colGroup = rowColGroup.add('group');
    colGroup.orientation = 'row';
    colGroup.add('statictext', undefined, LABELS.columns[lang]);
    var columnsInput = colGroup.add('edittext', undefined, "6");
    columnsInput.characters = 3;
    var rowGroup = rowColGroup.add('group');
    rowGroup.orientation = 'row';
    rowGroup.add('statictext', undefined, LABELS.rows[lang]);
    var rowsInput = rowGroup.add('edittext', undefined, "4");
    rowsInput.characters = 3;

    /* 2カラム部 / Two-column area */
    var columnsWrap = dlg.add('group');
    columnsWrap.orientation = 'row';
    columnsWrap.alignChildren = 'top';

    /* 左カラム: ピース数 / 形状 / オーバーラップ */
    var leftCol = columnsWrap.add('group');
    leftCol.orientation = 'column';
    leftCol.alignChildren = 'left';
    leftCol.margins = [10, 5, 10, 5];

    /* ピース数 */
    var totalPiecesGroup = leftCol.add("group");
    totalPiecesGroup.orientation = "row";
    totalPiecesGroup.alignment = "left";
    var totalPiecesLabel = totalPiecesGroup.add("statictext", undefined, LABELS.totalPieces[lang]);
    var totalPiecesInput = totalPiecesGroup.add("edittext", undefined, "25");
    totalPiecesInput.characters = 4;

    function getSelectedArtworkItem() {
      if (app.documents.length > 0 && app.selection.length == 1) {
        var selectedItem = app.selection[0];
        if (
          selectedItem.typename === "RasterItem" ||
          selectedItem.typename === "PlacedItem" ||
          selectedItem.typename === "SymbolItem" ||
          selectedItem.typename === "PathItem" ||
          selectedItem.typename === "GroupItem" ||
          selectedItem.typename === "CompoundPathItem"
        ) {
          var gb = selectedItem.geometricBounds;
          var width = gb[2] - gb[0];
          var height = gb[1] - gb[3];
          if (height < 0) height = -height;
          return { width: width, height: height };
        }
      }
      return null;
    }
    var selectedImage = getSelectedArtworkItem();
    if (selectedImage) {
      var grid = getInitialGridSize(selectedImage.width, selectedImage.height, 25);
      rowsInput.text = String(grid[0]);
      columnsInput.text = String(grid[1]);
    }

    function autoCalcRowsCols() {
      if (app.documents.length > 0 && app.selection.length === 1) {
        var selectedItem = app.activeDocument.selection[0];
        if (
          selectedItem.typename === "PlacedItem" ||
          selectedItem.typename === "SymbolItem" ||
          selectedItem.typename === "PathItem" ||
          selectedItem.typename === "GroupItem" ||
          selectedItem.typename === "CompoundPathItem"
        ) {
          if (!isNaN(parseInt(totalPiecesInput.text, 10)) && parseInt(totalPiecesInput.text, 10) > 0) {
            var pieceCount = parseInt(totalPiecesInput.text, 10);
          } else {
            return;
          }
          var bounds = selectedItem.geometricBounds;
          var w = bounds[2] - bounds[0];
          var h = bounds[1] - bounds[3];
          if (h < 0) h = -h;
          var aspect = w / h;
          var cols = Math.round(Math.sqrt(pieceCount * aspect));
          if (cols < 1) cols = 1;
          var rows = Math.round(pieceCount / cols);
          if (rows < 1) rows = 1;
          rowsInput.text = rows;
          columnsInput.text = cols;
        }
      }
    }

    totalPiecesInput.onChanging = function () {
      var val = parseInt(totalPiecesInput.text, 10);
      if (isNaN(val) || val < 1) return;
      var selectedImage = getSelectedArtworkItem();
      if (selectedImage) {
        var grid = getInitialGridSize(selectedImage.width, selectedImage.height, val);
        rowsInput.text = String(grid[0]);
        columnsInput.text = String(grid[1]);
      }
      autoCalcRowsCols();
    };

    columnsInput.onChanging = function () { };
    rowsInput.onChanging = function () { };

    /* 形状 */
    var shapeGroup = leftCol.add("group");
    shapeGroup.orientation = "column";
    shapeGroup.alignChildren = "left";
    var shapeRadioTraditional = shapeGroup.add("radiobutton", undefined, LABELS.shapeTraditional[lang]);
    var shapeRadioRandom = shapeGroup.add("radiobutton", undefined, LABELS.shapeRandom[lang]);
    shapeRadioTraditional.value = true;

    /* オーバーラップ */
    var overlapGroup = leftCol.add("group");
    overlapGroup.orientation = "column";
    overlapGroup.alignChildren = "left";
    var overlapCheckbox = overlapGroup.add('checkbox', undefined, LABELS.overlap[lang]);
    overlapCheckbox.value = true;
    var overlapInputRow = overlapGroup.add("group");
    overlapInputRow.orientation = "row";
    overlapInputRow.alignChildren = "left";
    overlapInputRow.margins = [20, 0, 0, 0];
    var overlapInput = overlapInputRow.add("edittext", undefined, "0");
    overlapInput.characters = 4;
    var overlapUnitLabel = overlapInputRow.add("statictext", undefined, getCurrentUnitLabel());
    overlapInput.enabled = overlapCheckbox.value;
    overlapUnitLabel.enabled = overlapCheckbox.value;
    overlapCheckbox.onClick = function () {
      var gridMode = modeRadioGrid.value;
      overlapInput.enabled = gridMode && overlapCheckbox.value;
      overlapUnitLabel.enabled = gridMode && overlapCheckbox.value;
    };

    /* 右カラム: オフセット / バラけ処理 / ケイ / 角丸 */
    var rightCol = columnsWrap.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = 'left';
    rightCol.margins = [10, 5, 10, 5];

    /* オフセット */
    var offsetGroup = rightCol.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = "left";
    var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
    offsetCheckbox.value = true;
    var offsetValueInput = offsetGroup.add("edittext", undefined, "-2");
    offsetValueInput.characters = 4;
    var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetValueInput.enabled = offsetCheckbox.value;
    offsetUnitLabel.enabled = offsetCheckbox.value;
    offsetCheckbox.onClick = function () {
      var puzzle = modeRadioPuzzle.value;
      offsetValueInput.enabled = puzzle && offsetCheckbox.value;
      offsetUnitLabel.enabled = puzzle && offsetCheckbox.value;
    };

    /* バラけ処理 */
    var scatterGroup = rightCol.add("group");
    scatterGroup.orientation = "row";
    scatterGroup.alignChildren = "left";
    var scatterCheckbox = scatterGroup.add('checkbox', undefined, LABELS.explode[lang]);
    scatterCheckbox.value = false;
    var scatterStrengthInput = scatterGroup.add("edittext", undefined, "30");
    scatterStrengthInput.characters = 4;
    scatterStrengthInput.enabled = scatterCheckbox.value;
    scatterCheckbox.onClick = function () {
      scatterStrengthInput.enabled = scatterCheckbox.value;
    };

    /* ケイ */
    var ruleGroup = rightCol.add("group");
    ruleGroup.orientation = "row";
    ruleGroup.alignChildren = "left";
    var ruleCheckbox = ruleGroup.add('checkbox', undefined, LABELS.ruleCheck[lang]);
    ruleCheckbox.value = false;

    /* 角丸 */
    var roundCornerGroup = rightCol.add("group");
    roundCornerGroup.orientation = "row";
    roundCornerGroup.alignChildren = "left";
    var roundCornerCheckbox = roundCornerGroup.add('checkbox', undefined, LABELS.roundCheck[lang]);
    roundCornerCheckbox.value = false;
    var roundRadiusInput = roundCornerGroup.add("edittext", undefined, "3");
    roundRadiusInput.characters = 5;
    var roundCornerUnitLabel = roundCornerGroup.add("statictext", undefined, getCurrentUnitLabel());
    roundRadiusInput.enabled = roundCornerCheckbox.value;
    roundCornerUnitLabel.enabled = roundCornerCheckbox.value;
    roundCornerCheckbox.onClick = function () {
      var gridMode = modeRadioGrid.value;
      roundRadiusInput.enabled = gridMode && roundCornerCheckbox.value;
      roundCornerUnitLabel.enabled = gridMode && roundCornerCheckbox.value;
    };

    /* モードに応じてディム化 / Toggle enabled state by mode */
    function updateModeDependentEnabled() {
      var puzzle = modeRadioPuzzle.value;
      /* パズル時のみ有効 / Puzzle only */
      shapeGroup.enabled = puzzle;
      totalPiecesInput.enabled = puzzle;
      totalPiecesLabel.enabled = puzzle;
      offsetCheckbox.enabled = puzzle;
      offsetValueInput.enabled = puzzle && offsetCheckbox.value;
      offsetUnitLabel.enabled = puzzle && offsetCheckbox.value;
      scatterCheckbox.enabled = puzzle;
      scatterStrengthInput.enabled = puzzle && scatterCheckbox.value;
      /* グリッド分割時のみ有効 / Grid only */
      overlapCheckbox.enabled = !puzzle;
      overlapInput.enabled = !puzzle && overlapCheckbox.value;
      overlapUnitLabel.enabled = !puzzle && overlapCheckbox.value;
      ruleCheckbox.enabled = !puzzle;
      roundCornerCheckbox.enabled = !puzzle;
      roundRadiusInput.enabled = !puzzle && roundCornerCheckbox.value;
      roundCornerUnitLabel.enabled = !puzzle && roundCornerCheckbox.value;
    }

    /* モード切替時の既定値適用 / Apply defaults when mode changes */
    function applyModeDefaults() {
      if (modeRadioGrid.value) {
        totalPiecesInput.text = "2";
        columnsInput.text = "2";
        rowsInput.text = "1";
        offsetValueInput.text = "0";
        offsetCheckbox.value = false;
        overlapCheckbox.value = true;
      } else {
        totalPiecesInput.text = "25";
        columnsInput.text = "6";
        rowsInput.text = "4";
        offsetValueInput.text = "-2";
        offsetCheckbox.value = true;
        overlapCheckbox.value = false;
      }
      overlapInput.text = "0";
      ruleCheckbox.value = false;
      roundCornerCheckbox.value = false;
    }

    function onModeChange() {
      applyModeDefaults();
      updateModeDependentEnabled();
    }
    modeRadioPuzzle.onClick = onModeChange;
    modeRadioGrid.onClick = onModeChange;
    /* デフォルトはグリッド分割 / Default to grid split */
    applyModeDefaults();
    updateModeDependentEnabled();

    /* OK・キャンセルボタン / OK and Cancel buttons */
    var buttonGroup = dlg.add('group');
    buttonGroup.orientation = 'row';
    buttonGroup.alignment = "center";
    buttonGroup.add('button', undefined, LABELS.cancel[lang], { name: "cancel" });
    var okBtn = buttonGroup.add('button', undefined, LABELS.okBtn[lang], { name: "ok" });
    okBtn.active = true;

    // Add arrow-key increment/decrement support for edittext fields
    changeValueByArrowKey(columnsInput, false);
    changeValueByArrowKey(rowsInput, false);
    changeValueByArrowKey(totalPiecesInput, false);
    changeValueByArrowKey(overlapInput, false);
    changeValueByArrowKey(roundRadiusInput, false);
    changeValueByArrowKey(offsetValueInput, true);
    changeValueByArrowKey(scatterStrengthInput, false);

    return {
      dialog: dlg,
      modeRadioPuzzle: modeRadioPuzzle,
      modeRadioGrid: modeRadioGrid,
      totalPiecesInput: totalPiecesInput,
      columnsInput: columnsInput,
      rowsInput: rowsInput,
      shapeRadioTraditional: shapeRadioTraditional,
      shapeRadioRandom: shapeRadioRandom,
      scatterCheckbox: scatterCheckbox,
      scatterStrengthInput: scatterStrengthInput,
      offsetCheckbox: offsetCheckbox,
      offsetValueInput: offsetValueInput,
      overlapCheckbox: overlapCheckbox,
      overlapInput: overlapInput,
      ruleCheckbox: ruleCheckbox,
      roundCornerCheckbox: roundCornerCheckbox,
      roundRadiusInput: roundRadiusInput
    };
  }

  /* オフセットパスエフェクトユーティリティ / Offset Path Effect Utility */
  function createOffsetEffectXML(offsetVal) {
    var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 2 "/></LiveEffect>';
    return xml.replace("value", offsetVal);
  }

  function applyOffsetPathToSelection(offsetVal) {
    try {
      var doc = app.activeDocument;
      var sourceItem = doc.selection[0];

      app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
      doc.selection = null;

      var duplicatedItem = sourceItem.duplicate(sourceItem, ElementPlacement.PLACEAFTER);
      sourceItem.remove();

      duplicatedItem.selected = true;
      duplicatedItem.applyEffect(createOffsetEffectXML(offsetVal));
      app.redraw();
      app.executeMenuCommand('expandStyle');

      app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

      if (app.selection && app.selection.length > 0) {
        return app.selection[0];
      } else {
        return null;
      }
    } catch (err) {
      alert("エラーが発生しました：\n" + err.message);
      return null;
    }
  }

  /* ケイ・角丸を適用 / Apply stroke (rule) and round corners to target */
  function applyRuleAndRoundCorners(target, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints) {
    if (!target) return;
    if (shouldAddStroke) {
      try {
        app.activeDocument.selection = null;
        target.selected = true;
        app.executeMenuCommand('Adobe New Stroke Shortcut');
        app.executeMenuCommand('Live Pathfinder Add');
      } catch (e) { }
    }
    if (shouldApplyRoundCorners && roundRadiusInPoints > 0) {
      try {
        var roundXML = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + roundRadiusInPoints + ' "/></LiveEffect>';
        target.applyEffect(roundXML);
      } catch (e) { }
    }
  }

  /* 矩形の座標をオフセットする関数 / Function to offset rectangle bounds */
  function offsetRectangleBounds(pathItem, offsetVal) {
    var gb = pathItem.geometricBounds;
    var newWidth = (gb[2] - gb[0]) + offsetVal * 2;
    var newHeight = (gb[1] - gb[3]) + offsetVal * 2;
    pathItem.left = pathItem.left - offsetVal;
    pathItem.top = pathItem.top + offsetVal;
    pathItem.width = newWidth;
    pathItem.height = newHeight;
    return pathItem;
  }

  function readDialogValues(ui) {
    var ruleCheckbox = ui.ruleCheckbox;
    var roundCornerCheckbox = ui.roundCornerCheckbox;
    var roundRadiusInput = ui.roundRadiusInput;
    var overlapCheckbox = ui.overlapCheckbox;
    var overlapInput = ui.overlapInput;
    var columnsInput = ui.columnsInput;
    var rowsInput = ui.rowsInput;
    var scatterCheckbox = ui.scatterCheckbox;
    var scatterStrengthInput = ui.scatterStrengthInput;
    var offsetCheckbox = ui.offsetCheckbox;
    var offsetValueInput = ui.offsetValueInput;

    var shouldAddStroke = ruleCheckbox.value;
    var shouldApplyRoundCorners = roundCornerCheckbox.value;
    var roundRadiusInputValue = parseFloat(roundRadiusInput.text);
    var roundRadiusInPoints = (isNaN(roundRadiusInputValue) ? 0 : roundRadiusInputValue) * getUnitToPtFactor();

    var overlapInPoints = 0;
    if (overlapCheckbox.value) {
      var overlapInputValue = parseFloat(overlapInput.text);
      overlapInPoints = (isNaN(overlapInputValue) ? 0 : overlapInputValue) * getUnitToPtFactor();
    }

    var shouldScatter = scatterCheckbox.value;
    var scatterStrengthInputValue = parseFloat(scatterStrengthInput.text);
    var scatterStrength = isNaN(scatterStrengthInputValue) ? 0 : scatterStrengthInputValue;

    var shouldApplyOffset = offsetCheckbox.value;
    var offsetInputValue = parseFloat(offsetValueInput.text);
    var offsetInPoints = (isNaN(offsetInputValue) ? 0 : offsetInputValue) * getUnitToPtFactor();

    var columnCount = Math.round(Number(columnsInput.text));
    var rowCount = Math.round(Number(rowsInput.text));

    return {
      shouldAddStroke: shouldAddStroke,
      shouldApplyRoundCorners: shouldApplyRoundCorners,
      roundRadiusInPoints: roundRadiusInPoints,
      overlapInPoints: overlapInPoints,
      shouldScatter: shouldScatter,
      scatterStrength: scatterStrength,
      shouldApplyOffset: shouldApplyOffset,
      offsetInPoints: offsetInPoints,
      columnCount: columnCount,
      rowCount: rowCount
    };
  }

  function prepareSourceItems() {
    var sourceItem;
    var contentSourceItem = null;
    var maskSourceItem;
    var isTemporaryBoundsRect = false;

    /* 複数選択時はグループ化してシンボル化（重ね順を保持） */
    if (app.selection.length > 1) {
      try {
        var tempGroup = app.activeDocument.groupItems.add();
        var selectedItems = [];
        for (var i = 0; i < app.selection.length; i++) {
          selectedItems.push(app.selection[i]);
        }
        for (var j = selectedItems.length - 1; j >= 0; j--) {
          selectedItems[j].move(tempGroup, ElementPlacement.PLACEATBEGINNING);
        }
        var groupLeft = tempGroup.left;
        var groupTop = tempGroup.top;
        var groupParent = tempGroup.parent;
        var groupedSymbol = app.activeDocument.symbols.add(tempGroup);
        var symbolItem = groupParent.symbolItems.add(groupedSymbol);
        symbolItem.left = groupLeft;
        symbolItem.top = groupTop;
        tempGroup.remove();
        sourceItem = symbolItem;
      } catch (e) {
        alert("複数オブジェクトのシンボル化に失敗しました: " + e);
        return null;
      }
    } else {
      sourceItem = app.selection[0];
    }
    maskSourceItem = sourceItem;

    /* 埋め込み画像(RasterItem)をSymbolItemに変換 */
    if (sourceItem.typename === "RasterItem") {
      try {
        var rasterGB = sourceItem.geometricBounds;
        var rasterLeft = sourceItem.left;
        var rasterTop = sourceItem.top;
        var rasterParent = sourceItem.parent;
        var tempSymbol = app.activeDocument.symbols.add(sourceItem);
        var rasterSymbolItem = rasterParent.symbolItems.add(tempSymbol);
        rasterSymbolItem.left = rasterLeft;
        rasterSymbolItem.top = rasterTop;
        sourceItem.remove();
        sourceItem = rasterSymbolItem;
      } catch (e) {
        alert("埋め込み画像のシンボル化に失敗しました: " + e);
        return null;
      }
    }
    /* ベクターオブジェクトをSymbolItemに変換 */
    else if (
      sourceItem.typename === "PathItem" ||
      sourceItem.typename === "GroupItem" ||
      sourceItem.typename === "CompoundPathItem"
    ) {
      try {
        var vectorLeft = sourceItem.left;
        var vectorTop = sourceItem.top;
        var vectorParent = sourceItem.parent;
        var vectorSymbol = app.activeDocument.symbols.add(sourceItem);
        var vectorSymbolItem = vectorParent.symbolItems.add(vectorSymbol);
        vectorSymbolItem.left = vectorLeft;
        vectorSymbolItem.top = vectorTop;
        sourceItem.remove();
        sourceItem = vectorSymbolItem;
      } catch (e) {
        alert("ベクターオブジェクトのシンボル化に失敗しました: " + e);
        return null;
      }
    }

    /* PlacedItem/RasterItem/SymbolItemは矩形化 */
    var sourceType = sourceItem.typename;
    if (sourceType == "PlacedItem" || sourceType == "RasterItem" || sourceType == "SymbolItem") {
      var gb = sourceItem.geometricBounds;
      var rectWidth = gb[2] - gb[0];
      var rectHeight = gb[1] - gb[3];
      if (rectHeight < 0) rectHeight = -rectHeight;
      var temporaryBoundsRect = app.activeDocument.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);
      maskSourceItem = temporaryBoundsRect;
      contentSourceItem = sourceItem;
      isTemporaryBoundsRect = true;
    }

    return {
      sourceItem: sourceItem,
      contentSourceItem: contentSourceItem,
      maskSourceItem: maskSourceItem,
      isTemporaryBoundsRect: isTemporaryBoundsRect
    };
  }

  function createClippedPiece(contentSourceItem, maskPath) {
    var placedCopy = contentSourceItem.duplicate();
    var clippingGroup = app.activeDocument.groupItems.add();
    placedCopy.moveToBeginning(clippingGroup);
    if (maskPath && maskPath.typename === "PathItem") {
      maskPath.moveToBeginning(clippingGroup);
      maskPath.clipping = true;
      clippingGroup.clipped = true;
    } else {
      alert("マスク用オブジェクトが PathItem ではないため、マスクをスキップします。");
    }
    return clippingGroup;
  }

  function finalizePieceAppearance(pieceItem, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints) {
    if (shouldScatter) {
      var offsetX = Math.random() * scatterStrength * 2 - scatterStrength;
      var offsetY = Math.random() * scatterStrength * 2 - scatterStrength;
      var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
      pieceItem.transform(moveMatrix);
    }
    pieceItem.selected = true;
    applyRuleAndRoundCorners(pieceItem, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
  }

  function cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect) {
    if (contentSourceItem && typeof contentSourceItem.remove === "function") {
      contentSourceItem.remove();
    }
    if (isTemporaryBoundsRect && maskSourceItem && typeof maskSourceItem.remove === "function") {
      maskSourceItem.remove();
    }
  }

  function createGridMask(originX, originY, columnIndex, rowIndex, pieceWidth, pieceHeight, overlapInPoints, imgLeft, imgRight, imgTop, imgBottom) {
    /* オーバーラップ考慮の矩形計算 / Rect with overlap, clamped to image bounds */
    var rectLeft = originX + columnIndex * pieceWidth - overlapInPoints / 2;
    var rectWidthActual = pieceWidth + overlapInPoints;
    if (rectLeft < imgLeft) {
      rectWidthActual -= (imgLeft - rectLeft);
      rectLeft = imgLeft;
    }
    if (rectLeft + rectWidthActual > imgRight) {
      rectWidthActual = imgRight - rectLeft;
    }

    var rectTop = originY - rowIndex * pieceHeight + overlapInPoints / 2;
    var rectHeightActual = pieceHeight + overlapInPoints;
    if (rectTop > imgTop) {
      rectHeightActual -= (rectTop - imgTop);
      rectTop = imgTop;
    }
    if (rectTop - rectHeightActual < imgBottom) {
      rectHeightActual = rectTop - imgBottom;
    }

    var gridMask = app.activeDocument.pathItems.rectangle(
      rectTop,
      rectLeft,
      rectWidthActual,
      rectHeightActual
    );
    gridMask.closed = true;
    gridMask.filled = false;
    gridMask.stroked = false;
    return gridMask;
  }

  function addPoint(obj, x, y) {
    var np = obj.pathPoints.add();
    np.anchor = [x, y];
    np.leftDirection = [x, y];
    np.rightDirection = [x, y];
    np.pointType = PointType.CORNER;
  }

  function createPuzzleMaskPath(originX, originY, x, y, pieceWidth, pieceHeight, columnCount, rowCount, oneThirdWide, oneQuarterHigh, oneThirdHigh, oneQuarterWide, pieceEdgeData) {

    var maskPath = app.activeDocument.pathItems.add();

    addPoint(maskPath, originX + x * pieceWidth, originY - (y + 1) * pieceHeight);

    appendTopEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, rowCount, oneThirdWide, oneQuarterHigh, pieceEdgeData);

    addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - (y + 1) * pieceHeight);

    appendRightEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, columnCount, oneThirdHigh, oneQuarterWide, pieceEdgeData);

    addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - y * pieceHeight);

    appendBottomEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdWide, oneQuarterHigh, pieceEdgeData);

    addPoint(maskPath, originX + x * pieceWidth, originY - y * pieceHeight);

    appendLeftEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdHigh, oneQuarterWide, pieceEdgeData);

    maskPath.closed = true;

    return maskPath;

  }

  function main() {
    app.undoGroup = "ジグソー作成";
    try {

      var ui = createDialog();
      var puzzlifyDialog = ui.dialog;
      var modeRadioGrid = ui.modeRadioGrid;
      var totalPiecesInput = ui.totalPiecesInput;
      var columnsInput = ui.columnsInput;
      var rowsInput = ui.rowsInput;
      var shapeRadioRandom = ui.shapeRadioRandom;
      var scatterCheckbox = ui.scatterCheckbox;
      var scatterStrengthInput = ui.scatterStrengthInput;
      var offsetCheckbox = ui.offsetCheckbox;
      var offsetValueInput = ui.offsetValueInput;
      var overlapCheckbox = ui.overlapCheckbox;
      var overlapInput = ui.overlapInput;
      var ruleCheckbox = ui.ruleCheckbox;
      var roundCornerCheckbox = ui.roundCornerCheckbox;
      var roundRadiusInput = ui.roundRadiusInput;

      /* Illustrator状態とダイアログ入力チェック / Illustrator state and dialog input check */
      var selectedImage = null;
      if (
        app.documents.length > 0 &&
        app.selection.length > 0 &&
        puzzlifyDialog.show() == 1 &&
        (columnsInput.text || rowsInput.text)
      ) {
        /* ダイアログ値取得 / Read dialog values */
        var dialogValues = readDialogValues(ui);
        var shouldAddStroke = dialogValues.shouldAddStroke;
        var shouldApplyRoundCorners = dialogValues.shouldApplyRoundCorners;
        var roundRadiusInPoints = dialogValues.roundRadiusInPoints;
        var overlapInPoints = dialogValues.overlapInPoints;
        var shouldScatter = dialogValues.shouldScatter;
        var scatterStrength = dialogValues.scatterStrength;
        var shouldApplyOffset = dialogValues.shouldApplyOffset;
        var offsetInPoints = dialogValues.offsetInPoints;

        // 選択オブジェクト取得 / Get selected object
        var preparedItems = prepareSourceItems();
        if (!preparedItems) {
          return;
        }
        var sourceItem = preparedItems.sourceItem;
        var contentSourceItem = preparedItems.contentSourceItem;
        var maskSourceItem = preparedItems.maskSourceItem;
        var isTemporaryBoundsRect = preparedItems.isTemporaryBoundsRect;
        selectedImage = app.selection[0];

        /* 列・行数取得 */
        var columnCount = dialogValues.columnCount;
        var rowCount = dialogValues.rowCount;

        if (
          (columnCount >= 1 && rowCount >= 1) ||
          (columnCount == 0 && rowCount >= 1) ||
          (columnCount >= 1 && rowCount == 0)
        ) {
          maskSourceItem.selected = false;

          var aspectRatio;
          if (columnCount == 0) {
            aspectRatio = Math.abs(
              (maskSourceItem.geometricBounds[2] - maskSourceItem.geometricBounds[0]) /
              (maskSourceItem.geometricBounds[3] - maskSourceItem.geometricBounds[1])
            );
            columnCount = Math.round(aspectRatio * rowCount);
          }
          if (rowCount == 0) {
            aspectRatio = Math.abs(
              (maskSourceItem.geometricBounds[3] - maskSourceItem.geometricBounds[1]) /
              (maskSourceItem.geometricBounds[2] - maskSourceItem.geometricBounds[0])
            );
            rowCount = Math.round(aspectRatio * columnCount);
          }

          var bounds = maskSourceItem.geometricBounds;
          var left = bounds[0];
          var top = bounds[1];
          var right = bounds[2];
          var bottom = bounds[3];

          var originX = left;
          var originY = top;

          var pieceWidth = (right - left) / columnCount;
          var pieceHeight = (top - bottom) / rowCount;

          var oneThirdWide = pieceWidth / 3;
          var oneQuarterHigh = pieceHeight / 4;
          var oneThirdHigh = pieceHeight / 3;
          var oneQuarterWide = pieceWidth / 4;

          var pieceEdgeData = new Array(rowCount);
          if (shapeRadioRandom.value) {
            for (var y = 0; y < rowCount; y++) {
              pieceEdgeData[y] = new Array(columnCount);
              for (var x = 0; x < columnCount; x++) {
                pieceEdgeData[y][x] = {
                  topOut: Math.random() < 0.5,
                  rightOut: Math.random() < 0.5,
                  verticalOffset: pieceHeight * (Math.random() - 0.5) / 10,
                  horizontalOffset: pieceWidth * (Math.random() - 0.5) / 10
                };
              }
            }
          } else {
            for (var y = 0; y < rowCount; y++) {
              pieceEdgeData[y] = new Array(columnCount);
              for (var x = 0; x < columnCount; x++) {
                pieceEdgeData[y][x] = {
                  topOut: (x & 1) ^ (y & 1),
                  rightOut: !((x & 1) ^ (y & 1)),
                  verticalOffset: pieceHeight * (Math.random() - 0.5) / 10,
                  horizontalOffset: pieceWidth * (Math.random() - 0.5) / 10
                };
              }
            }
          }

          /* グリッド形状（矩形マスク）モード */
          if (modeRadioGrid.value) {
            var imgLeft = left;
            var imgRight = right;
            var imgTop = top;
            var imgBottom = bottom;
            var generatedPieces = [];
            for (var y = 0; y < rowCount; y++) {
              for (var x = 0; x < columnCount; x++) {
                var gridMask = createGridMask(
                  originX,
                  originY,
                  x,
                  y,
                  pieceWidth,
                  pieceHeight,
                  overlapInPoints,
                  imgLeft,
                  imgRight,
                  imgTop,
                  imgBottom
                );
                if (contentSourceItem && (contentSourceItem.typename === "PlacedItem" || contentSourceItem.typename === "SymbolItem")) {
                  var clippedPiece = createClippedPiece(contentSourceItem, gridMask);
                  generatedPieces.push(clippedPiece);
                  finalizePieceAppearance(clippedPiece, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
                } else {
                  generatedPieces.push(gridMask);
                  finalizePieceAppearance(gridMask, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
                }
              }
            }

            cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect);
            return;
          }

          /* ジグソー形状モード */
          var generatedPieces = [];
          for (var y = 0; y < rowCount; y++) {
            for (var x = 0; x < columnCount; x++) {
              var maskPath = createPuzzleMaskPath(
                originX,
                originY,
                x,
                y,
                pieceWidth,
                pieceHeight,
                columnCount,
                rowCount,
                oneThirdWide,
                oneQuarterHigh,
                oneThirdHigh,
                oneQuarterWide,
                pieceEdgeData
              );

              /* オフセット効果を適用 */
              if (shouldApplyOffset && maskPath) {
                var offsetVal = offsetInPoints;
                app.selection = null;
                try {
                  maskPath.selected = true;
                  var offsetResultItem = applyOffsetPathToSelection(offsetVal);
                  if (offsetResultItem) {
                    if (offsetResultItem.typename === "PathItem") {
                      maskPath = offsetResultItem;
                    } else if (offsetResultItem.typename === "GroupItem") {
                      for (var i = 0; i < offsetResultItem.pageItems.length; i++) {
                        if (offsetResultItem.pageItems[i].typename === "PathItem") {
                          maskPath = offsetResultItem.pageItems[i];
                          break;
                        }
                      }
                      if (maskPath.typename !== "PathItem") {
                        alert("オフセット後の GroupItem に PathItem が含まれていません。マスク処理をスキップします。");
                      }
                    } else {
                      alert("オフセット後のオブジェクトが予期しない型です。マスク処理をスキップします。");
                    }
                  }
                } catch (e) {
                  alert("オフセット適用中にエラーが発生しました：" + e.message);
                }
              }

              /* 画像クリッピンググループ処理 */
              if (contentSourceItem && (contentSourceItem.typename === "PlacedItem" || contentSourceItem.typename === "SymbolItem")) {
                var clippedPiece = createClippedPiece(contentSourceItem, maskPath);
                generatedPieces.push(clippedPiece);
                finalizePieceAppearance(clippedPiece, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
              } else {
                generatedPieces.push(maskPath);
                finalizePieceAppearance(maskPath, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
              }
            }
          }

          cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect);
        }
      } else {
        puzzlifyDialog.hide();
      }


    } catch (e) {
      alert("スクリプト実行中にエラーが発生しました: " + e);
    }
  }

  main();

})();