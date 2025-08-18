#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
レイヤー統合（フラット化）を行うIllustrator用スクリプト。
除外名（例：bg/背景/background）を持つレイヤーを残しつつ、その他のレイヤー配下の全オブジェクトを
「_mergedLayer」に集約（移動）します。その後、空になったレイヤーを再帰的に検出して削除します。
*/

### スクリプト名：

レイヤー統合（フラット化） / Flatten Layers

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### 概要：

- 除外レイヤーを残し、それ以外のレイヤー配下のオブジェクトを「_mergedLayer」に移動
- 統合後、空になった最上位レイヤーおよびサブレイヤーを再帰的に削除

### 主な機能：

- レイヤー配下オブジェクトの一括移動（ロック一時解除→移動→再ロック）
- 除外レイヤー名（bg / 背景 / background）対応
- 空レイヤーの再帰削除（UI なし／alert なし）

### 処理の流れ：

1. ドキュメント取得（未オープンなら終了）  
2. 既存「_mergedLayer」を取得（なければ作成）  
3. 除外レイヤーを除いて、全レイヤー配下のオブジェクトを「_mergedLayer」に移動  
4. ドキュメント内の空レイヤーを再帰的に収集し、一括削除

### 更新履歴：

- v1.0 (20250414) : 初期バージョン  
- v1.1 (20250818) : 微調整  
- v1.2 (20250818) : 空レイヤー再帰削除ロジックを組み込み、説明文を再構成

---

/*
Illustrator script to flatten layers. It consolidates all objects (except in excluded layers
such as bg/背景/background) into a single layer named "_mergedLayer", then recursively
removes layers that became empty.
*/

### Script Name:

Flatten Layers

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### Overview:

- Move all objects from non-excluded layers into "_mergedLayer"  
- Recursively delete empty top-level and sublayers after consolidation

### Key Features:

- Bulk move items (temporarily unlock → move → restore lock)  
- Excluded layer names (bg / 背景 / background)  
- Recursive empty-layer deletion (no UI / no alerts)

### Process Flow:

1. Get active document (exit if none)  
2. Ensure or create "_mergedLayer"  
3. Move items from all non-excluded layers into "_mergedLayer"  
4. Collect and delete empty layers recursively

### Update History:

- v1.0 (20250414) : Initial release  
- v1.1 (20250818) : Minor adjustments  
- v1.2 (20250818) : Added recursive empty-layer removal and rewrote the description

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

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

    // --- 空レイヤーの再帰収集と削除 / Collect and delete empty layers recursively ---
    var emptyLayerList = [];
    findEmptyLayers(documentRef, emptyLayerList);
    deleteLayers(emptyLayerList);
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

// ==========================
// 空レイヤーを収集する関数（再帰処理） / Collect empty layers recursively
// ==========================
function findEmptyLayers(container, resultArray) {
    var layers = container.layers;
    for (var i = 0; i < layers.length; i++) {
        var currentLayer = layers[i];

        // 子レイヤーを先にチェック / Recurse into sublayers first
        if (currentLayer.layers.length > 0) {
            findEmptyLayers(currentLayer, resultArray);
        }

        // 直下にページアイテムが存在せず、子レイヤーも空の場合 → 削除対象
        var hasItems = currentLayer.pageItems.length > 0;
        var hasChildLayers = currentLayer.layers.length > 0;
        if (!hasItems && !hasChildLayers) {
            resultArray.push(currentLayer);
        }
    }
}

// ==========================
// 指定レイヤー配列を削除する関数 / Delete listed layers
// ==========================
function deleteLayers(layerArray) {
    for (var i = 0; i < layerArray.length; i++) {
        try { layerArray[i].remove(); } catch (e) {}
    }
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