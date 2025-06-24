#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：SendToBgLayer.jsx
概要：選択されたオブジェクトを「bg」レイヤーに重ね順を維持したまま移動し、最背面に配置します。
処理の流れ：
  1. ドキュメントと現在のアクティブレイヤーを取得
  2. 「bg」レイヤーの存在を確認し、なければ作成
  3. 選択されたオブジェクトを重ね順を維持したまま「bg」レイヤーに移動（エラーを無視）
  4. 「bg」レイヤーを最背面に移動してロック
  5. 元のレイヤーを再アクティブ化
対象：選択されたオブジェクト
除外：未選択時には何もしない
更新日：2025-06-24
*/

function main() {
    var activeDoc = app.activeDocument;
    var originalLayer = activeDoc.activeLayer; // 元のアクティブレイヤーを記憶

    // 「bg」レイヤーが存在するか確認し、なければ作成
    var bgLayer;
    try {
        bgLayer = activeDoc.layers.getByName("bg");
        bgLayer.locked = false; // ロックを解除
    } catch (e) {
        bgLayer = activeDoc.layers.add(); // 新規作成
        bgLayer.name = "bg";
    }

    var selectedItems;
    try {
        selectedItems = activeDoc.selection; // 現在の選択オブジェクトを取得
    } catch (e) {
        selectedItems = [];
    }

    if (selectedItems && selectedItems.length > 0) {
        // 選択オブジェクトを重ね順を維持したまま「bg」レイヤーに移動
        for (var i = 0; i < selectedItems.length; i++) {
            try {
                selectedItems[i].move(bgLayer, ElementPlacement.PLACEATEND);
            } catch (e) {
                // エラーがあっても処理を続行
            }
        }
    }

    bgLayer.zOrder(ZOrderMethod.SENDTOBACK); // 「bg」レイヤー自体を最背面に
    bgLayer.locked = true; // 「bg」レイヤーを再ロック

    // 処理終了後、元のレイヤーを再アクティブ化
    activeDoc.activeLayer = originalLayer;
}

main();