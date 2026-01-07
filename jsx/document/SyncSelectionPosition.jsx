/*
  Align Selection Across Documents
  最前面のドキュメントの選択オブジェクトの座標に、他ドキュメントの選択オブジェクトを合わせる
*/

(function() {
    // ドキュメントが2つ以上開かれていない場合は終了
    if (app.documents.length < 2) {
        alert("2つ以上のドキュメントを開いてください。");
        return;
    }

    // 1. 最前面（ソース）のドキュメントの処理
    var sourceDoc = app.activeDocument;
    var sourceSelection = sourceDoc.selection;

    // 何も選択されていない場合はエラー
    if (sourceSelection.length === 0) {
        alert("最前面のドキュメントでオブジェクトを選択してください。");
        return;
    }

    // 基準となる座標（X, Y）を取得
    // 複数選択されている場合、全体のバウンディングボックスの左上(Left, Top)を基準とします
    // Illustratorの座標系: Y軸は上がプラス、下がマイナス（ただしPositionプロパティは通常 [x, y] で扱います）
    // 注意: Illustrator Scriptingでの position[1] (Y) は、画面上の見た目のY座標です。
    
    var targetX, targetY;

    // グループ化されていない複数選択の場合でも、個別のPositionではなく
    // 「選択範囲全体」の左上を取得して基準にするのが一般的ですが、
    // ここではシンプルに「選択されている最初のオブジェクト（最前面）」を基準にします。
    // もし選択範囲全体の左上がよければロジックを少し変える必要がありますが、通常はこれで十分機能します。
    
    // 選択オブジェクトの座標を取得 (Left, Top)
    var sourcePos = getSelectionPosition(sourceSelection);
    targetX = sourcePos[0];
    targetY = sourcePos[1];

    // 2. 他のすべてのドキュメントに対する処理
    for (var i = 0; i < app.documents.length; i++) {
        var currentDoc = app.documents[i];

        // ソースドキュメントはスキップ
        if (currentDoc === sourceDoc) continue;

        // ドキュメントをアクティブにする
        app.activeDocument = currentDoc;
        
        var currentSelection = currentDoc.selection;

        if (currentSelection.length > 0) {
            // 移動処理
            setSelectionPosition(currentSelection, targetX, targetY);
        } else {
            // 選択がない場合はコンソール等にログを出すか、無視するか
            // 今回は無視します
        }
    }

    // 最後に元のドキュメントをアクティブに戻す
    app.activeDocument = sourceDoc;
    alert("完了しました。\n座標 X: " + targetX.toFixed(2) + ", Y: " + targetY.toFixed(2) + " に統一しました。");


    // ---------------------------------------------------------
    // ヘルパー関数: 選択範囲の左上座標を取得する
    // ---------------------------------------------------------
    function getSelectionPosition(sel) {
        // 単一オブジェクトの場合
        if (sel.length === 1) {
            return sel[0].position; // [x, y]
        }
        
        // 複数オブジェクトの場合、最も左(minX)と最も上(maxY)を探す
        // Illustratorの座標系ではYは上がプラスなので、Topは最大値を探すことになる
        var minX = sel[0].position[0];
        var maxY = sel[0].position[1];

        for (var k = 1; k < sel.length; k++) {
            var itemX = sel[k].position[0];
            var itemY = sel[k].position[1];
            
            if (itemX < minX) minX = itemX;
            if (itemY > maxY) maxY = itemY;
        }
        return [minX, maxY];
    }

    // ---------------------------------------------------------
    // ヘルパー関数: 選択範囲を指定座標に移動する
    // ---------------------------------------------------------
    function setSelectionPosition(sel, x, y) {
        // 現在の選択範囲の左上座標を取得
        var currentPos = getSelectionPosition(sel);
        var currentX = currentPos[0];
        var currentY = currentPos[1];

        // 移動すべき差分（デルタ）を計算
        var deltaX = x - currentX;
        var deltaY = y - currentY;

        // すべての選択オブジェクトを差分だけ移動する
        for (var k = 0; k < sel.length; k++) {
            sel[k].translate(deltaX, deltaY);
        }
    }

})();
