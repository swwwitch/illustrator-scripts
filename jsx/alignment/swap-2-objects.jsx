#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    スクリプト名: SwapObjectsByCenter.jsx
    概要:
    - Illustratorで2つのオブジェクトを選択しているとき、
      それぞれの中心位置を入れ替えるスクリプト。
    - 通常オブジェクトは visibleBounds、クリップグループはマスクパスの geometricBounds を基準にする。
    - 2つ以外の選択数では警告を表示。

    更新履歴:
    - v1.0 (20250802) : 初版作成
    - v1.1 (20250802) : クリップグループのマスクパスを基準に対応
*/

function swapObjectsByCenter(objA, objB) {
    // 中心点を取得する関数
    function getCenter(obj) {
        var bounds;
        if (obj.typename === "GroupItem" && obj.clipped) {
            // クリップグループの場合はマスクパスを基準に
            var mask = null;
            for (var i = 0; i < obj.pageItems.length; i++) {
                if (obj.pageItems[i].clipping) {
                    mask = obj.pageItems[i];
                    break;
                }
            }
            if (mask) {
                bounds = mask.geometricBounds; // マスクパスの範囲
            } else {
                bounds = obj.visibleBounds; // 保険：マスクが見つからなければ全体
            }
        } else {
            bounds = obj.visibleBounds;
        }
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    // 中心座標を取得
    var centerA = getCenter(objA);
    var centerB = getCenter(objB);

    // 移動量を計算
    var deltaA = [centerB[0] - centerA[0], centerB[1] - centerA[1]];
    var deltaB = [centerA[0] - centerB[0], centerA[1] - centerB[1]];

    // 位置を入れ替え
    objA.translate(deltaA[0], deltaA[1]);
    objB.translate(deltaB[0], deltaB[1]);
}

function main() {
    try {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var sel = app.activeDocument.selection;
        if (!sel || sel.length !== 2) {
            alert("2つのオブジェクトを選択してください。");
            return;
        }

        var objA = sel[0];
        var objB = sel[1];

        swapObjectsByCenter(objA, objB);

    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

main();