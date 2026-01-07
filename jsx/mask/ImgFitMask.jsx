(function() {
    // ドキュメントが開かれていない場合は終了
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // 選択アイテムが2つでない場合はエラー
    if (sel.length !== 2) {
        alert("エラー: 1つの図形（マスク用）と1つの画像を選択してください。");
        return;
    }

    var maskObj = null; // マスクになる図形
    var targetImg = null; // マスクされる画像

    // 選択アイテムを分類する
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
            maskObj = item;
        } else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {
            targetImg = item;
        }
    }

    // 適切なオブジェクトが見つからなかった場合
    if (!maskObj || !targetImg) {
        alert("エラー: 「パス（図形）」と「配置画像」をそれぞれ1つずつ選択してください。");
        return;
    }

    // --- 1. スケール調整（隙間が出ないようにリサイズ） ---

    var maskW = maskObj.width;
    var maskH = maskObj.height;
    var imgW = targetImg.width;
    var imgH = targetImg.height;

    // 幅と高さ、それぞれの倍率を計算
    var scaleX = maskW / imgW;
    var scaleY = maskH / imgH;

    // 隙間が出ないようにするには、倍率が大きい方を採用する (Math.max)
    // ※画像全体を収める（余白が出てもいい）場合は Math.min にします
    var scaleFactor = Math.max(scaleX, scaleY);

    // リサイズ実行（Illustratorのresizeはパーセント指定なので100倍する）
    // 引数: scaleX, scaleY, changePositions, changeFillPatterns, changeFillGradients, changeStrokePattern
    targetImg.resize(scaleFactor * 100, scaleFactor * 100, true, true, true, true);


    // --- 2. 中央揃えの処理 ---
    
    // 図形の中心座標
    var maskCenter = [
        maskObj.left + maskObj.width / 2,
        maskObj.top - maskObj.height / 2
    ];

    // 画像の中心座標（リサイズ後のサイズで計算）
    var newImgCenter = [
        targetImg.left + targetImg.width / 2,
        targetImg.top - targetImg.height / 2
    ];

    // 移動距離を計算
    var deltaX = maskCenter[0] - newImgCenter[0];
    var deltaY = maskCenter[1] - newImgCenter[1];

    // 画像を移動
    targetImg.translate(deltaX, deltaY);


    // --- 3. クリッピングマスクの作成処理 ---

    var clipGroup = doc.groupItems.add();
    
    // 元の画像があった位置（重ね順）の付近にグループを移動
    clipGroup.move(maskObj, ElementPlacement.PLACEBEFORE);

    // オブジェクトをグループ内に移動（マスク用パスを最前面に）
    maskObj.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
    targetImg.move(clipGroup, ElementPlacement.PLACEATEND);

    // クリップ設定
    clipGroup.clipped = true;

    // 選択状態の更新
    doc.selection = null;
    clipGroup.selected = true;

})();
