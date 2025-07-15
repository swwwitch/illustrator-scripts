#target illustrator
$.localize = true;

/*
### スクリプトの説明 / Script Description

このスクリプトはAdobe Illustratorのドキュメント内に存在する「_guide」という名前のレイヤーを対象としています。
対象レイヤーのロックを解除し、そのレイヤー内のすべてのガイドを解除します。
解除したガイドは「UnlockedGuides」という名前の新しいレイヤーに移動され、移動後に再度ロックされます。

This script targets a layer named "_guide" in the active Adobe Illustrator document.
It unlocks the target layer, removes the guide attribute from all items within it,
then moves those items to a new layer named "UnlockedGuides", which is locked after the move.

### 更新履歴 / Change Log

v1.0 (20250716) : 初期バージョン / Initial version
*/

var SCRIPT_VERSION = "v1.0";
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

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
        guideLayer.locked = false; /* ロックを解除 / Unlock the layer */

        // 「UnlockedGuides」レイヤーを探す
        var newLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name == "UnlockedGuides") {
                newLayer = doc.layers[i];
                // ロック解除が必要なら解除
                if (newLayer.locked) {
                    newLayer.locked = false; /* ロックを解除 / Unlock the layer */
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
                items[j].guides = false; /* ガイドを解除 / Remove guide flag */
            }
            /* 外観設定：塗りなし、線はK100、1pt / Set appearance: no fill, stroke K100, 1pt */
            items[j].filled = false;
            items[j].stroked = true;
            items[j].strokeColor = new GrayColor();
            items[j].strokeColor.gray = 100;
            items[j].strokeWidth = 1;
            // 「UnlockedGuides」レイヤーに移動
            items[j].move(newLayer, ElementPlacement.PLACEATBEGINNING);
        }

        guideLayer.locked = true; /* 再ロック / Relock the layer */
    }
}

function main() {
    unlockGuideLayerAndClearGuides();
}

main();