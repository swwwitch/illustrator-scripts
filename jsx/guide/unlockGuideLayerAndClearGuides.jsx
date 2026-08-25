#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「_guide」レイヤーのロックを解除し、そのレイヤー内のすべてのガイドを解除します。
解除したオブジェクトは「UnlockedGuides」レイヤーへ移動し、移動後にそのレイヤーを再ロックします。

詳細は README を参照してください。

### Overview

Unlocks the "_guide" layer and removes the guide attribute from every item in it.
The released items are moved to a new "UnlockedGuides" layer, which is locked once the move is done.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "unlockGuideLayerAndClearGuides"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/unlockGuideLayerAndClearGuides.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/unlockGuideLayerAndClearGuides.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 「_guide」レイヤーがあれば、そのロックを解除して、ガイドを解除後、「UnlockedGuides」レイヤーに移動し再ロックする / If a "_guide" layer exists, unlock it, remove guides, move to "UnlockedGuides" layer, and relock */
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

        /* 「UnlockedGuides」レイヤーを探す / Find "UnlockedGuides" layer */
        var newLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name == "UnlockedGuides") {
                newLayer = doc.layers[i];
                /* ロック解除が必要なら解除 / Unlock if necessary */
                if (newLayer.locked) {
                    newLayer.locked = false; /* ロックを解除 / Unlock the layer */
                }
                break;
            }
        }
        /* 見つからなければ作成 / Create if not found */
        if (!newLayer) {
            newLayer = doc.layers.add();
            newLayer.name = "UnlockedGuides";
        }

        var items = guideLayer.pageItems;
        for (var j = items.length - 1; j >= 0; j--) {
            /* ガイドフラグがあれば解除 / Remove guide flag if present */
            if (items[j].guides) {
                items[j].guides = false; /* ガイドを解除 / Remove guide flag */
            }
            /* 外観設定：塗りなし、線はK100、1pt / Set appearance: no fill, stroke K100, 1pt */
            items[j].filled = false;
            items[j].stroked = true;
            items[j].strokeColor = new GrayColor();
            items[j].strokeColor.gray = 100;
            items[j].strokeWidth = 1;
            /* 「UnlockedGuides」レイヤーに移動 / Move to "UnlockedGuides" layer */
            items[j].move(newLayer, ElementPlacement.PLACEATBEGINNING);
        }

        guideLayer.locked = true; /* 再ロック / Relock the layer */
    }
}

function main() {
    unlockGuideLayerAndClearGuides();
}

main();
