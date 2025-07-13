// Illustrator用のJavaScript

// 「_guide」レイヤーがあれば、そのロックを解除して、ガイドを解除後、「UnlockedGuides」レイヤーに移動し再ロックする
function unlockGuideLayerAndClearGuides() {
    var doc = app.activeDocument;
    var guideLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name == "_guide") {
            guideLayer = doc.layers[i];
            break;
        }
    }
    if (guideLayer) {
        guideLayer.locked = false; // ロックを解除

        // 「UnlockedGuides」レイヤーを探す
        var newLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name == "UnlockedGuides") {
                newLayer = doc.layers[i];
                // ロック解除が必要なら解除
                if (newLayer.locked) {
                    newLayer.locked = false;
                }
                break;
            }
        }
        // 見つからなければ作成
        if (!newLayer) {
            newLayer = doc.layers.add();
            newLayer.name = "UnlockedGuides";
        }

        var items = guideLayer.pageItems;
        for (var j = items.length - 1; j >= 0; j--) {
            // ガイドフラグがあれば解除
            if (items[j].guides) {
                items[j].guides = false; // ガイドを解除
            }
            // 外観設定：塗りなし、線はK100、1pt
            items[j].filled = false;
            items[j].stroked = true;
            items[j].strokeColor = new GrayColor();
            items[j].strokeColor.gray = 100;
            items[j].strokeWidth = 1;
            // 「UnlockedGuides」レイヤーに移動
            items[j].move(newLayer, ElementPlacement.PLACEATBEGINNING);
        }

        guideLayer.locked = true; // 再ロック
    }
}

function main() {
    unlockGuideLayerAndClearGuides();
}

main();