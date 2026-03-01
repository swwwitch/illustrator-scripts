(function() {
    // ドキュメントが開かれていない場合は終了
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length < 2) {
        alert("エラー: 複数の画像と図形を選択してください。");
        return;
    }

    var images = [];
    var masks = [];

    // 1. 選択アイテムを「画像」と「図形」に分類
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
            // クリップグループ内のパスなどが誤って選択されないよう、親がLayerかGroupのみ対象にするなど
            // 簡易的なチェックですが、通常選択ならこれでOK
            masks.push(item);
        } else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {
            images.push(item);
        }
    }

    if (images.length === 0 || masks.length === 0) {
        alert("画像と図形がそれぞれ少なくとも1つずつ必要です。");
        return;
    }

    // 距離計算用の関数
    function getDistance(p1, p2) {
        return Math.sqrt(Math.pow(p2[0] - p1[0], 2) + Math.pow(p2[1] - p1[1], 2));
    }

    // 中心座標取得用の関数
    function getCenter(item) {
        return [
            item.left + item.width / 2,
            item.top - item.height / 2
        ];
    }

    // 2. マッチングと処理の実行
    // 画像を基準に、一番近いマスクを探すループ
    // ※処理済みマスクを重複使用しないように管理する
    var usedMasks = [];

    // マッチング精度を高めるため、左上(X,Y座標)順などでソートしても良いですが、
    // 今回は「総当たりで最短距離ペア」を見つける方式にします。

    // ペアリストを作成 {img: item, mask: item, dist: number}
    var pairs = [];

    for (var i = 0; i < images.length; i++) {
        var img = images[i];
        var imgCenter = getCenter(img);
        var bestMask = null;
        var minDist = Infinity;

        for (var j = 0; j < masks.length; j++) {
            var msk = masks[j];
            var dist = getDistance(imgCenter, getCenter(msk));
            
            if (dist < minDist) {
                minDist = dist;
                bestMask = msk;
            }
        }
        
        if (bestMask) {
            pairs.push({
                image: img,
                mask: bestMask,
                distance: minDist
            });
        }
    }

    // 距離が近い順にソート（これが重要：遠くの誤判定を防ぐ）
    pairs.sort(function(a, b) {
        return a.distance - b.distance;
    });

    // ペアごとに処理実行（マスクが重複しないようにチェック）
    var processedCount = 0;
    
    for (var k = 0; k < pairs.length; k++) {
        var p = pairs[k];
        
        // このマスクがまだ使われていなければ処理実行
        var isUsed = false;
        for(var u=0; u<usedMasks.length; u++){
            if(usedMasks[u] === p.mask) {
                isUsed = true;
                break;
            }
        }

        if (!isUsed) {
            processClip(p.image, p.mask);
            usedMasks.push(p.mask);
            processedCount++;
        }
    }

    // 完了メッセージ（任意）
    // alert(processedCount + " 組のマスクを作成しました。");


    // --- 個別のマスク処理関数 ---
    function processClip(targetImg, maskObj) {
        
        // スケール調整（隙間が出ないようにリサイズ）
        var maskW = maskObj.width;
        var maskH = maskObj.height;
        var imgW = targetImg.width;
        var imgH = targetImg.height;

        var scaleX = maskW / imgW;
        var scaleY = maskH / imgH;
        var scaleFactor = Math.max(scaleX, scaleY);

        targetImg.resize(scaleFactor * 100, scaleFactor * 100, true, true, true, true);

        // 中央揃えの処理
        var maskCenter = getCenter(maskObj);
        var newImgCenter = getCenter(targetImg); // リサイズ後の中心再取得

        var deltaX = maskCenter[0] - newImgCenter[0];
        var deltaY = maskCenter[1] - newImgCenter[1];

        targetImg.translate(deltaX, deltaY);

        // クリッピングマスクの作成
        var clipGroup = doc.groupItems.add();
        clipGroup.move(maskObj, ElementPlacement.PLACEBEFORE);
        maskObj.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        targetImg.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;
    }

})();
