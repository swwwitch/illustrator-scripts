#target illustrator

    /*
    概要：
    選択中のオブジェクトがグループ内にある場合、親グループを辿って所属レイヤー直下へ移動します。
    重ね順の逆転を避けるため、選択オブジェクトは逆順（後ろから）に処理します。
    
    更新日：2026-03-05
    */

    (function () {

        function moveSelectedObjectsToLayer() {
            if (app.documents.length === 0) {
                alert("ドキュメントが開かれていません。");
                return;
            }

            var doc = app.activeDocument;
            var sel = doc.selection;

            // 1. 選択されているオブジェクトがあるか確認
            if (sel.length === 0) {
                alert("オブジェクトを選択して実行してください。");
                return;
            }

            // 2. 選択されたすべてのオブジェクトを退避
            var targets = [];
            for (var i = 0; i < sel.length; i++) {
                targets.push(sel[i]);
            }

            // 3. 重ね順が逆転しないように「逆順（後ろから）」で処理を行う
            for (var j = targets.length - 1; j >= 0; j--) {
                var obj = targets[j];
                var parentObj = obj.parent;

                // すでにレイヤー直下にいる（グループに入っていない）場合は何もしない
                if (parentObj.typename !== "GroupItem") {
                    continue;
                }

                // 親が「レイヤー」になるまで、階層を上へ上へと遡る
                var targetLayer = parentObj;
                while (targetLayer.typename === "GroupItem") {
                    targetLayer = targetLayer.parent;
                }

                // 見つかったレイヤーにオブジェクトを移動
                if (targetLayer.typename === "Layer") {
                    // レイヤーの最前面に移動
                    // 逆順で処理しているため、結果的に元の上下関係が維持される
                    obj.move(targetLayer, ElementPlacement.PLACEATBEGINNING);

                    // 移動後も選択状態を維持
                    obj.selected = true;
                }
            }
        }

        moveSelectedObjectsToLayer();

    })();