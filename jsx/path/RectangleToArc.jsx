#target illustrator

/*
 * RectangleToArc.jsx
 *
 * 概要 / Overview:
 *   選択した長方形を、左下隅・上辺中央・右下隅の 3 点を通る円弧に変換する。
 *   長方形の幅と高さから円の半径と中心を求め、その円弧をベジェ曲線
 *   （最大 90 度ごとのセグメント）で生成する。
 *   元の長方形は削除し、線の色・太さは生成した円弧に引き継ぐ。
 *
 *   Converts each selected rectangle into a circular arc passing through
 *   its bottom-left corner, top-edge midpoint, and bottom-right corner.
 *
 * 対象 / Target: パスアイテム（長方形）。複数選択に対応。
 */

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    var SCRIPT_VERSION = "v1.0.0";

    function main() {
        // ドキュメントが開かれているか確認
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。");
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        // オブジェクトが選択されているか確認
        if (selectedItems.length === 0) {
            alert("長方形を選択してから実行してください。");
            return;
        }

        // 削除処理を行うため、選択アイテムを配列にコピーしておく
        var itemsToProcess = [];
        for (var i = 0; i < selectedItems.length; i++) {
            itemsToProcess.push(selectedItems[i]);
        }

        for (var i = 0; i < itemsToProcess.length; i++) {
            var sourceRect = itemsToProcess[i];

            // パスアイテム以外はスキップ
            if (sourceRect.typename !== "PathItem") {
                continue;
            }

            // バウンディングボックスを取得
            var rectBounds = sourceRect.geometricBounds;
            var left = rectBounds[0];
            var top = rectBounds[1];
            var right = rectBounds[2];
            var bottom = rectBounds[3];

            // 幅と高さを計算
            var rectWidth = Math.abs(right - left);
            var rectHeight = Math.abs(top - bottom);
            if (rectWidth === 0 || rectHeight === 0) continue;

            // 円の半径 arcRadius と中心点 (centerX, centerY) を計算
            var arcRadius = (rectWidth * rectWidth) / (8 * rectHeight) + (rectHeight / 2);
            var centerX = left + rectWidth / 2;
            var centerY = top - arcRadius;

            // 左下端と右下端の角度を計算
            var dy = bottom - centerY;

            // 左下のポイントの角度 (開始角)
            var startAngle = Math.atan2(dy, left - centerX);
            // 上端を通るように角度を調整（90度 = PI/2 より大きくする）
            while (startAngle < Math.PI / 2) {
                startAngle += 2 * Math.PI;
            }

            // 右下のポイントの角度 (終了角)
            var endAngle = Math.atan2(dy, right - centerX);
            // 上端を通るように角度を調整（90度 = PI/2 より小さくする）
            while (endAngle > Math.PI / 2) {
                endAngle -= 2 * Math.PI;
            }

            // 元の長方形と同じ階層・位置に新しいパス（弧）を作成
            var arcPath = sourceRect.layer.pathItems.add();
            arcPath.move(sourceRect, ElementPlacement.PLACEBEFORE);

            // ベジェ曲線で正確な円弧を描くため、最大90度ごとにセグメントを分割する
            var numSegments = Math.ceil((startAngle - endAngle) / (Math.PI / 2));
            var angleStep = (startAngle - endAngle) / numSegments;

            for (var j = 0; j <= numSegments; j++) {
                var currentAngle = startAngle - j * angleStep;
                var pathPoint = arcPath.pathPoints.add();

                // アンカーポイントの座標
                var anchorX = centerX + arcRadius * Math.cos(currentAngle);
                var anchorY = centerY + arcRadius * Math.sin(currentAngle);
                pathPoint.anchor = [anchorX, anchorY];

                // ベジェ曲線のハンドル長さを計算する公式
                var handleLength = (4 / 3) * arcRadius * Math.tan(angleStep / 4);

                // 時計回りの接線ベクトル
                var tangentX = Math.sin(currentAngle);
                var tangentY = -Math.cos(currentAngle);

                // ハンドルの方向を設定
                pathPoint.leftDirection = [anchorX - handleLength * tangentX, anchorY - handleLength * tangentY];
                pathPoint.rightDirection = [anchorX + handleLength * tangentX, anchorY + handleLength * tangentY];
            }

            // 弧の見た目（塗りなし・線あり）を設定
            arcPath.filled = false;
            arcPath.stroked = true;

            // 元の長方形の線の色・太さを引き継ぐ
            if (sourceRect.stroked) {
                arcPath.strokeColor = sourceRect.strokeColor;
                arcPath.strokeWidth = sourceRect.strokeWidth;
            }

            // 元の長方形を削除
            sourceRect.remove();
        }
    }

    main();

})();