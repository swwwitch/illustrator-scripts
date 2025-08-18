#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
アートボードサイズを調整するIllustrator用スクリプト。
選択範囲やアートボード全体を対象に、ユーザー入力に基づいてサイズを調整できます。
プレビュー機能付きの処理を備え、数値入力を直感的に操作可能です。
*/

### スクリプト名：

アートボードサイズ調整スクリプト

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- Illustrator内のアートボードサイズを数値入力で調整
- プレビューを確認しながら操作可能

### 主な機能：

- アートボードサイズの数値指定による調整
- プレビュー境界チェック機能
- 設定の即時反映

### 処理の流れ：

1. 数値を入力してプレビュー確認  
2. 確定後、アートボードサイズを変更  

### 更新履歴：

- v1.0 (20250414) : 初期バージョン
- v1.1 (20250818) : 微調整

---

/*
Illustrator script for adjusting artboard size.
Allows resizing based on user input with preview support.
*/

### Script Name:

Adjust Artboard Size Script

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Adjusts Illustrator artboard size via numeric input  
- Provides preview functionality for intuitive control  

### Key Features:

- Resize artboard by entering dimensions  
- Toggle preview boundaries  
- Real-time update of preview  

### Process Flow:

1. Enter values and check preview  
2. Apply resizing after confirmation  

### Update History:

- v1.0 (20250414) : Initial release
- v1.1 (20250818) : 

---

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

function main() {
    var documentRef;
    try {
        documentRef = app.activeDocument;
    } catch (error) {
        return;
    }

    var allTopLayers = documentRef.layers;
    var mergedLayerName = "_mergedLayer";

    var mergedLayer = null;
    try {
        mergedLayer = documentRef.layers.getByName(mergedLayerName);
    } catch (e) {
        mergedLayer = null;
    }
    if (mergedLayer === null) {
        mergedLayer = documentRef.layers.add();
        mergedLayer.name = mergedLayerName;
    }

    mergedLayer.visible = true;
    mergedLayer.locked = false;
    mergedLayer.printable = true;

    var customRGBColor = new RGBColor();
    customRGBColor.red = 79;
    customRGBColor.green = 127;
    customRGBColor.blue = 255;
    mergedLayer.color = customRGBColor;


    var topCount = allTopLayers.length;
    for (var i = 0; i < topCount; i++) {
        var currentLayer = allTopLayers[i];

        if (isExcludedLayer(currentLayer.name) || currentLayer === mergedLayer) {
            continue;
        }

        moveItemsToTargetLayer(currentLayer, mergedLayer);
    }

    for (var j = allTopLayers.length - 1; j >= 0; j--) {
        var layerToCheck = allTopLayers[j];

        if (isExcludedLayer(layerToCheck.name) || layerToCheck === mergedLayer) {
            continue;
        }

        if (isLayerEmpty(layerToCheck)) {
            try {
                layerToCheck.remove();
            } catch (error) {}
        }
    }
}

/*
 * 指定レイヤー配下の全ページアイテムを再帰的に destinationLayer へ移動。
 * Move all page items under sourceLayer recursively to destinationLayer.
 * Returns true if any items were moved.
 */
function moveItemsToTargetLayer(sourceLayer, destinationLayer) {
    if (sourceLayer === destinationLayer) return false;
    if (sourceLayer.locked || !sourceLayer.visible) return false;
    var didMove = false;

    var sublayers = sourceLayer.layers;
    for (var i = 0, n = sublayers.length; i < n; i++) {
        if (moveItemsToTargetLayer(sublayers[i], destinationLayer)) {
            didMove = true;
        }
    }

    var pageItems = sourceLayer.pageItems;
    // ロック項目は一時解除→移動→元に戻す / Temporarily unlock, move, then restore lock
    for (var j = pageItems.length - 1; j >= 0; j--) {
        var item = pageItems[j];
        var wasLocked = false;
        try {
            wasLocked = item.locked;
            if (wasLocked) item.locked = false; // temporarily unlock
            item.move(destinationLayer, ElementPlacement.PLACEATEND);
            didMove = true;
        } catch (error) {
            // 移動できない項目は無視 / ignore items that still cannot be moved
        } finally {
            try {
                if (wasLocked) item.locked = true; // restore lock state
            } catch (e2) {}
        }
    }

    return didMove;
}

/* Exclude-name map for layers (case-insensitive, trimmed) */
var EXCLUDE = {
    "bg": 1,
    "背景": 1,
    "background": 1
};

/* Check if the layer name is excluded from merging (trim + case-insensitive) */
function isExcludedLayer(layerName) {
    if (!layerName) return false;
    var key = String(layerName).replace(/^\s+|\s+$/g, '').toLowerCase();
    return EXCLUDE[key] === 1;
}

/*
 * レイヤーが“実質空”かを厳密に判定。
 * Judge if a layer is effectively empty (no sublayers and no items in any collection).
 */
function isLayerEmpty(layer) {
    // sublayers exist → not empty
    if (layer.layers.length > 0) return false;

    // any items in known collections → not empty
    if (layer.pageItems.length > 0) return false; // umbrella collection
    if (layer.groupItems.length > 0) return false;
    if (layer.pathItems.length > 0) return false;
    if (layer.compoundPathItems.length > 0) return false;
    if (layer.textFrames.length > 0) return false;
    if (layer.placedItems.length > 0) return false;
    if (layer.rasterItems.length > 0) return false;
    if (layer.graphItems && layer.graphItems.length > 0) return false;
    if (layer.meshItems && layer.meshItems.length > 0) return false;
    if (layer.symbolItems && layer.symbolItems.length > 0) return false;
    if (layer.pluginItems && layer.pluginItems.length > 0) return false;

    return true; // nothing found
}

main();