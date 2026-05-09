#target illustrator

(function () {
    // ===== 設定スイッチ =====
    var COLUMNS = 4;   // グリッドの列数
    var GAP = 100;     // アートボード間の余白（pt）
    // =======================

    function main() {
        // ドキュメントが開かれているかチェック
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;
        var artboards = doc.artboards;
        var artboardEntries = [];

        // 1. 現在のアートボードの情報を配列に保存（rect / name / 所属アイテム）
        for (var i = 0; i < artboards.length; i++) {
            artboardEntries.push({
                rect: artboards[i].artboardRect,
                name: artboards[i].name,
                items: []
            });
        }

        // シャッフル前に基準位置（元のアートボード 0 の左上）を記憶
        var anchorLeft = artboardEntries[0].rect[0];
        var anchorTop = artboardEntries[0].rect[1];

        // 2. 各ページアイテムを中心点で所属アートボードに割り当て（最初に一致したもの）
        assignItemsToArtboards(doc, artboardEntries);

        // 3. 配列をランダムにシャッフル
        shuffleArtboardOrder(artboardEntries);

        // 4. グリッドに再配置（中身のアイテムも一緒に移動）
        layoutArtboardsInGrid(artboardEntries, COLUMNS, GAP, anchorLeft, anchorTop);

        // 5. シャッフル＋再配置した情報を元のアートボードに上書き
        for (var i = 0; i < artboards.length; i++) {
            artboards[i].artboardRect = artboardEntries[i].rect;
            artboards[i].name = artboardEntries[i].name;
        }

        // 画面の更新（再描画）
        app.redraw();

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand("fitall");
    }

    // 配列をランダムにシャッフル（フィッシャー–イェーツのシャッフル）
    function shuffleArtboardOrder(artboardEntries) {
        for (var i = artboardEntries.length - 1; i > 0; i--) {
            var swapIndex = Math.floor(Math.random() * (i + 1));
            var temp = artboardEntries[i];
            artboardEntries[i] = artboardEntries[swapIndex];
            artboardEntries[swapIndex] = temp;
        }
    }

    // 各ページアイテムを中心点が含まれるアートボードに割り当てる
    function assignItemsToArtboards(doc, artboardEntries) {
        for (var itemIndex = 0; itemIndex < doc.pageItems.length; itemIndex++) {
            var item = doc.pageItems[itemIndex];
            // 入れ子のアイテムは親（グループ等）の移動で連動するためスキップ
            if (item.parent.typename !== "Layer") continue;
            if (item.locked || item.hidden) continue;
            if (item.parent.locked || item.parent.visible === false) continue;

            var bounds = item.geometricBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            var centerY = (bounds[1] + bounds[3]) / 2;

            for (var artboardIndex = 0; artboardIndex < artboardEntries.length; artboardIndex++) {
                var rect = artboardEntries[artboardIndex].rect;
                var minX = Math.min(rect[0], rect[2]);
                var maxX = Math.max(rect[0], rect[2]);
                var minY = Math.min(rect[1], rect[3]);
                var maxY = Math.max(rect[1], rect[3]);
                if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY) {
                    artboardEntries[artboardIndex].items.push(item);
                    break;
                }
            }
        }
    }

    // アートボードを N 列のグリッドに配置（所属アイテムも同じ移動量で平行移動）
    function layoutArtboardsInGrid(artboardEntries, columns, gap, anchorLeft, anchorTop) {
        if (artboardEntries.length === 0) return;

        // 各アートボードの最大幅・高さを取得してセルサイズを決定
        var maxWidth = 0, maxHeight = 0;
        for (var i = 0; i < artboardEntries.length; i++) {
            var rect = artboardEntries[i].rect;
            var width = rect[2] - rect[0];
            var height = Math.abs(rect[3] - rect[1]);
            if (width > maxWidth) maxWidth = width;
            if (height > maxHeight) maxHeight = height;
        }

        var cellWidth = maxWidth + gap;
        var cellHeight = maxHeight + gap;

        // 基準位置（シャッフル前の元のアートボード 0 の左上）
        var startLeft = anchorLeft;
        var startTop = anchorTop;

        // y 軸の向き（top と bottom の大小関係から判定）
        var referenceRect = artboardEntries[0].rect;
        var verticalDirection = (referenceRect[3] > referenceRect[1]) ? 1 : -1;

        for (var i = 0; i < artboardEntries.length; i++) {
            var col = i % columns;
            var row = Math.floor(i / columns);
            var oldRect = artboardEntries[i].rect;
            var artboardWidth = oldRect[2] - oldRect[0];
            var artboardHeight = oldRect[3] - oldRect[1];

            var newLeft = startLeft + col * cellWidth;
            var newTop = startTop + row * cellHeight * verticalDirection;

            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            // 所属アイテムを同じ移動量で平行移動
            var items = artboardEntries[i].items;
            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                try {
                    items[itemIndex].translate(deltaX, deltaY);
                } catch (e) {}
            }

            artboardEntries[i].rect = [newLeft, newTop, newLeft + artboardWidth, newTop + artboardHeight];
        }
    }

    main();
})();
