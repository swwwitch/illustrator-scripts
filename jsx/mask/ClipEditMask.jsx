/**
 * 2つの矩形（geometricBounds）が重なっているかを判定する
 * @param {number[]} boundsA - [left, top, right, bottom]（top > bottom）
 * @param {number[]} boundsB - [left, top, right, bottom]（top > bottom）
 * @returns {boolean} 重なっていれば true
 */
function isBoundsOverlapping(boundsA, boundsB) {
    var leftA = boundsA[0], topA = boundsA[1], rightA = boundsA[2], bottomA = boundsA[3];
    var leftB = boundsB[0], topB = boundsB[1], rightB = boundsB[2], bottomB = boundsB[3];

    var horizontalOverlap = (leftA < rightB) && (leftB < rightA);
    var verticalOverlap = (bottomA < topB) && (bottomB < topA);

    return horizontalOverlap && verticalOverlap;
}

/**
 * 選択オブジェクトを重なりで連結成分（クラスタ）に分割する
 * 直接重ならなくても、間のオブジェクトを介して繋がっていれば同じクラスタになる
 * @param {Array} items - PageItem の配列
 * @returns {Array} PageItem 配列の配列（クラスタごと）
 */
function clusterItemsByOverlap(items) {
    var bounds = [];
    for (var i = 0; i < items.length; i++) {
        bounds.push(items[i].geometricBounds);
    }

    var visited = [];
    for (var v = 0; v < items.length; v++) {
        visited.push(false);
    }

    var clusters = [];
    for (var start = 0; start < items.length; start++) {
        if (visited[start]) {
            continue;
        }

        // start を起点に、重なりで繋がるものを幅優先で集める
        var cluster = [];
        var queue = [start];
        visited[start] = true;

        while (queue.length > 0) {
            var current = queue.shift();
            cluster.push(items[current]);

            for (var other = 0; other < items.length; other++) {
                if (!visited[other] && isBoundsOverlapping(bounds[current], bounds[other])) {
                    visited[other] = true;
                    queue.push(other);
                }
            }
        }

        clusters.push(cluster);
    }

    return clusters;
}

/**
 * 指定した PageItem 群だけを選択状態にする
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - 選択したい PageItem の配列
 */
function selectOnly(doc, items) {
    doc.selection = null;
    for (var i = 0; i < items.length; i++) {
        items[i].selected = true;
    }
}

(function() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    // 何も選択されていない場合は処理を終了
    if (selection.length === 0) {
        return;
    }

    // 選択されているオブジェクトが1つの場合
    if (selection.length === 1) {
        var selectedItem = selection[0];

        // クリップグループ（GroupItemであり、かつclippedプロパティがtrue）の場合
        if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
            // マスクを編集 (Edit Mask)
            app.executeMenuCommand('editMask');
            app.executeMenuCommand('editMask');
        } else {
            // クリッピングマスクを作成 (Make Clipping Mask)
            app.executeMenuCommand('makeMask');
        }
        return;
    }

    // 選択されているオブジェクトが2つ以上の場合
    // 重なりで連結成分（クラスタ）に分けてから、クラスタ単位でマスクを作成する
    var selectedItems = [];
    for (var s = 0; s < selection.length; s++) {
        selectedItems.push(selection[s]);
    }

    var clusters = clusterItemsByOverlap(selectedItems);

    if (clusters.length === 1) {
        // すべてが1つの塊（重なっている）→ そのまま1回だけマスク
        app.executeMenuCommand('makeMask');
    } else {
        // 離れた塊が複数（重なっていない）→ 塊ごとにマスク
        for (var c = 0; c < clusters.length; c++) {
            var cluster = clusters[c];

            // 単独オブジェクトの塊で、それが画像（配置画像＝埋め込み / リンク画像）の場合
            if (cluster.length === 1 &&
                (cluster[0].typename === "RasterItem" || cluster[0].typename === "PlacedItem")) {
                // TODO: 配置画像・リンク画像なら…（処理を記述）
            }

            selectOnly(doc, cluster);
            app.executeMenuCommand('makeMask');
        }
    }
})();
