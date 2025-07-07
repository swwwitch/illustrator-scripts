#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddTrimMark.jsx

### 概要

- アートボードまたは選択オブジェクトにトリムマークを作成するIllustrator用スクリプトです。
- トリムマークは専用の「トンボ」レイヤーに配置され、オブジェクトのガイドも自動生成されます。

### 主な機能

- 選択オブジェクトがあればその形状を基にトリムマークを作成
- 選択がない場合はアートボード全体を基にトリムマークを作成
- トリムマークは「トンボ」レイヤーに移動し、自動ロック
- 元オブジェクトのガイド化と再配置

### 処理の流れ

1. 選択オブジェクトがあればそれを使用、なければアートボード矩形を作成
2. 対象オブジェクトを複製、塗り・線をなしに設定
3. トリムマーク作成メニュー実行後、複製オブジェクト削除
4. トリムマークを「トンボ」レイヤーに移動、レイヤーをロック
5. 元オブジェクトを複製してガイド化

### 更新履歴

- v1.0.0 (20250205) : 初期バージョン
- v1.0.1 (20250603) : コメント整理と処理安定化

---

### Script Name:

AddTrimMark.jsx

### Overview

- An Illustrator script to create trim marks for an artboard or selected objects.
- Trim marks are placed on a dedicated "Trim" layer, and a guide of the original object is also automatically created.

### Main Features

- Create trim marks based on selected object shape if available
- If no selection, use the entire artboard rectangle as base
- Move trim marks to a "Trim" layer and lock it automatically
- Duplicate the original object and convert to guide

### Process Flow

1. Use selected object if available, otherwise create rectangle from artboard
2. Duplicate target object, remove fill and stroke
3. Execute trim mark creation menu, then delete duplicate object
4. Move trim marks to "Trim" layer and lock it
5. Duplicate the original object and convert to guide

### Update History

- v1.0.0 (20250205): Initial version
- v1.0.1 (20250603): Refined comments and stabilized process
*/

function main() {
    var doc = app.activeDocument;
    var targetObj = null;

    if (doc.selection.length > 0) {
        // 選択オブジェクトがある場合
        targetObj = doc.selection[0];
    } else {
        // 選択オブジェクトがない場合、アクティブなアートボードを基に処理
        var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var rect = artboard.artboardRect; // [左, 上, 右, 下]

        targetObj = doc.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
        targetObj.filled = false;
        targetObj.stroked = false;
    }

    var duplicatedObj = targetObj.duplicate();

    // 塗りと線をなしにする
    duplicatedObj.filled = false;
    duplicatedObj.stroked = false;

    // 複製オブジェクトを選択状態にする
    doc.selection = [duplicatedObj];

    // トリムマークを作成
    app.executeMenuCommand('TrimMark v25');

    // 複製オブジェクトを削除
    duplicatedObj.remove();

    // 「トンボ」レイヤーを取得（なければ作成）
    var trimLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "トンボ") {
            trimLayer = doc.layers[i];
            break;
        }
    }
    if (!trimLayer) {
        trimLayer = doc.layers.add();
        trimLayer.name = "トンボ";
    }

    // 「トンボ」レイヤーをロック解除（移動処理のため）
    trimLayer.locked = false;

    // トリムマークを「トンボ」レイヤーに移動
    for (var j = 0; j < doc.selection.length; j++) {
        doc.selection[j].move(trimLayer, ElementPlacement.PLACEATBEGINNING);
    }

    // 「トンボ」レイヤーをロック
    trimLayer.locked = true;

    // 元のオブジェクトを再選択
    doc.selection = [targetObj];

    // 元のオブジェクトを複製し、ガイド化する
    var guideObj = targetObj.duplicate();
    guideObj.guides = true;
}

main();