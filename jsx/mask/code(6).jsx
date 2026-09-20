#target illustrator

// =========================================
// レイアウト設定 / Layout settings
// =========================================
var GAP_MM = 5; /* 切り詰めたあとの上下パーツの間隔（mm） */

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length !== 2) {
        alert("画像とパスを2つだけ選択してください。");
        return;
    }

    var image = null;
    var bandPath = null;
    var selTypes = [];

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        selTypes.push(item.typename);

        /* PlacedItem はリンク画像、RasterItem は埋め込み画像 */
        if (item.typename === "PlacedItem" || item.typename === "RasterItem") {
            image = item;
        } else if (
            item.typename === "PathItem" ||
            item.typename === "CompoundPathItem"
        ) {
            bandPath = item;
        }
    }

    if (!image) {
        alert("画像が選択されていません。\n選択中: " + selTypes.join(", "));
        return;
    }

    if (!bandPath) {
        alert("削除範囲を示すパスが選択されていません。");
        return;
    }

    /* 座標比較の許容値（pt） */
    var TOLERANCE = 0.001;

    /* geometricBounds は [left, top, right, bottom]、Y軸は上が大きい */
    var imageBounds = image.geometricBounds;
    var bandBounds = bandPath.geometricBounds;

    var imageLeft = imageBounds[0];
    var imageTop = imageBounds[1];
    var imageRight = imageBounds[2];
    var imageBottom = imageBounds[3];

    var bandTop = bandBounds[1];
    var bandBottom = bandBounds[3];

    if (bandTop >= imageTop - TOLERANCE || bandBottom <= imageBottom + TOLERANCE) {
        alert("削除する帯は、画像の上端・下端より内側に描いてください。");
        return;
    }

    var container = image.parent;
    var imageWidth = imageRight - imageLeft;

    /**
     * 画像を複製し、指定した横帯の矩形でクリッピングマスクを作成する
     * @param {number} rectTop - 矩形の上端のY座標
     * @param {number} rectHeight - 矩形の高さ
     * @return {GroupItem} 作成したクリップグループ
     */
    function createClipPart(rectTop, rectHeight) {
        var clipRect = container.pathItems.rectangle(rectTop, imageLeft, imageWidth, rectHeight);
        var partImage = image.duplicate(container, ElementPlacement.PLACEATBEGINNING);

        var clipGroup = container.groupItems.add();
        clipRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        partImage.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;

        return clipGroup;
    }

    /*
     * 帯の上端・下端を境に、残す範囲の矩形を2つ作って切り出す
     * 上側: 画像の上端 〜 帯の上端
     * 下側: 帯の下端 〜 画像の下端
     */
    var upperPart = createClipPart(imageTop, imageTop - bandTop);
    var lowerPart = createClipPart(bandBottom, bandBottom - imageBottom);

    /* 下側を上へ動かし、上側との間隔を GAP_MM にする */
    var gapPt = GAP_MM * 72 / 25.4;
    lowerPart.translate(0, bandTop - bandBottom - gapPt);

    /* 元画像と帯のパスは役目を終えるので削除 */
    image.remove();
    bandPath.remove();

    doc.selection = null;
    upperPart.selected = true;
    lowerPart.selected = true;
})();
