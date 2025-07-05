#target illustrator

/**
 * 【Illustrator用スクリプト】
 * アートボードまたは選択オブジェクトに対してトリムマークを作成し、
 * 「トンボ」レイヤーに配置してロックします。
 * あわせて対象オブジェクトを複製し、ガイドに変換します。
 *
 * 【処理の流れ】
 * 1. 選択オブジェクトがある場合はそれを使用。なければアートボードから矩形を作成
 * 2. 対象オブジェクトを複製し、塗りと線をなしに設定して選択状態に
 * 3. トリムマーク作成メニューを実行
 * 4. 複製オブジェクトを削除（トリムマークのみ残す）
 * 5. 「トンボ」レイヤーがなければ作成し、トリムマークを移動
 * 6. 「トンボ」レイヤーをロック
 * 7. 元のオブジェクトを複製してガイド化
 *
 * 【作成日】2025-02-05
 * 【更新日】2025-06-03
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