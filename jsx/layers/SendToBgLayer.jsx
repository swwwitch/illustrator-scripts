#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SendToBgLayer.jsx

### 概要

- 選択されたオブジェクトを「bg」レイヤーに重ね順を維持したまま移動し、最背面に配置するIllustrator用スクリプトです。
- 自動的に「bg」レイヤーが作成され、処理後にロックされます。

### 主な機能

- 選択オブジェクトを「bg」レイヤーに移動
- 重ね順を維持して配置
- 「bg」レイヤーを自動作成（存在しない場合）
- レイヤーの最背面配置とロック
- 元のアクティブレイヤーを復元

### 処理の流れ

1. ドキュメントとアクティブレイヤーを取得
2. 「bg」レイヤーの存在を確認、なければ作成
3. 選択オブジェクトを重ね順を維持して移動
4. 「bg」レイヤーを最背面に移動してロック
5. 元のレイヤーを再アクティブ化

### 更新履歴

- v1.0.0 (20240624) : 初期バージョン

---

### Script Name:

SendToBgLayer.jsx

### Overview

- An Illustrator script to move selected objects to a "bg" layer while preserving their stacking order and placing them at the very back.
- Automatically creates a "bg" layer if it does not exist, and locks it after processing.

### Main Features

- Move selected objects to "bg" layer
- Preserve stacking order
- Auto-create "bg" layer if missing
- Send "bg" layer to back and lock it
- Restore original active layer

### Process Flow

1. Get document and active layer
2. Check if "bg" layer exists, create if not
3. Move selected objects preserving stacking order
4. Send "bg" layer to back and lock
5. Reactivate original layer

### Update History

- v1.0.0 (20240624): Initial version
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