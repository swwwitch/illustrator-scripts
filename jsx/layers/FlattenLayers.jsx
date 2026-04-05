#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
レイヤー統合（フラット化）を行うIllustrator用スクリプト。
除外名（例：bg/背景/background）を持つレイヤーを残しつつ、その他のレイヤー配下の全オブジェクトを
「_mergedLayer」に集約（移動）します。その後、空になったレイヤーを再帰的に検出して削除します。

### スクリプト名：

レイヤー統合（フラット化） / Flatten Layers

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### 概要：

- 実行前にダイアログを表示し、処理条件を選択可能
- 除外レイヤーを残し、それ以外のレイヤー配下のオブジェクトを「_mergedLayer」に移動
- ロック / 非表示のレイヤー・オブジェクトをそれぞれ対象外にするか選択可能
- 空レイヤー削除、既存「_mergedLayer」の使用、レイヤーカラー設定の可否を選択可能

### 主な機能：

- ダイアログで処理条件を選択
- レイヤー配下オブジェクトの一括移動（ロック項目は一時解除→移動→再ロック）
- 除外レイヤー名（bg / 背景 / background）対応
- ロック / 非表示のレイヤー・オブジェクトの対象外切り替え
- 空レイヤーの再帰削除（ON/OFF）
- 既存「_mergedLayer」の使用 / 新規作成
- 「_mergedLayer」のレイヤーカラー設定（ON/OFF）

### 処理の流れ：

1. ドキュメント取得（未オープンなら終了）
2. ダイアログで処理条件を選択（キャンセル時は終了）
3. 設定に応じて既存「_mergedLayer」を取得、または新規作成
4. 除外レイヤーを除いて、条件に合う全レイヤー配下のオブジェクトを「_mergedLayer」に移動
5. 設定が ON の場合、ドキュメント内の空レイヤーを再帰的に収集し、一括削除

### 更新履歴：

- v1.0 (20250414) : 初期バージョン
- v1.1 (20250818) : 微調整
- v1.2 (20250818) : 空レイヤー再帰削除ロジックを組み込み、説明文を再構成
- v1.3 (20260406) : 前後関係維持ロジックを修正し、ロック／非表示レイヤーを対象外とする仕様を明記
- v1.4 (20260406) : ダイアログを追加し、ロック / 非表示のレイヤー・オブジェクト、空レイヤー削除、既存 mergedLayer 使用、レイヤーカラー設定を選択可能に

Illustrator script to flatten layers. It consolidates all objects (except in excluded layers
such as bg/背景/background) into a single layer named "_mergedLayer", then recursively
removes layers that became empty.

### Script Name:

Flatten Layers

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### Overview:

- Show a dialog before execution so the processing conditions can be selected
- Move objects from non-excluded layers into "_mergedLayer"
- Allow the user to choose whether locked / hidden layers and objects are excluded
- Allow the user to choose empty-layer deletion, reuse of an existing "_mergedLayer", and layer-color assignment

### Key Features:

- Select processing conditions in a dialog
- Bulk move items (temporarily unlock item → move → restore lock)
- Excluded layer names (bg / 背景 / background)
- Toggle exclusion of locked / hidden layers and objects
- Recursive empty-layer deletion (ON/OFF)
- Reuse existing "_mergedLayer" or create a new one
- Set "_mergedLayer" layer color (ON/OFF)

### Process Flow:

1. Get the active document (exit if none)
2. Show the options dialog (exit if canceled)
3. Reuse or create "_mergedLayer" according to the selected options
4. Move items from eligible non-excluded layers into "_mergedLayer"
5. If enabled, collect and delete empty layers recursively

### Update History:

- v1.0 (20250414) : Initial release
- v1.1 (20250818) : Minor adjustments
- v1.2 (20250818) : Added recursive empty-layer removal and rewrote the description
- v1.3 (20260406) : Fixed stacking-order preservation logic and documented that locked/hidden layers are excluded
- v1.4 (20260406) : Added a dialog so locked/hidden layers and objects, empty-layer deletion, merged-layer reuse, and layer-color assignment can be selected

*/

// スクリプトバージョン

var SCRIPT_VERSION = "v1.4";

function getCurrentLang() {
    return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: 'レイヤー統合（フラット化）',
        en: 'Flatten Layers'
    },
    options: {
        ja: '後処理',
        en: 'Cleanup'
    },
    exclude: {
        ja: '対象外にする',
        en: 'Exclude'
    },
    destination: {
        ja: 'まとめ先',
        en: 'Destination'
    },
    layerName: {
        ja: 'レイヤー名',
        en: 'Layer name'
    },
    layerColor: {
        ja: 'レイヤーカラー',
        en: 'Layer color'
    },
    skipLockedLayers: {
        ja: 'レイヤー',
        en: 'Layers'
    },
    skipHiddenLayers: {
        ja: 'レイヤー',
        en: 'Layers'
    },
    skipLockedObjects: {
        ja: 'オブジェクト',
        en: 'Objects'
    },
    skipHiddenObjects: {
        ja: 'オブジェクト',
        en: 'Objects'
    },
    deleteEmptyLayers: {
        ja: '空のレイヤー／サブレイヤーを削除',
        en: 'Delete empty layers / sublayers'
    },
    cancel: {
        ja: 'キャンセル',
        en: 'Cancel'
    },
    ok: {
        ja: 'OK',
        en: 'OK'
    }
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry.en;
}

function showOptionsDialog(hasExistingMergedLayer) {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';

    var destPanel = dlg.add('panel', undefined, L('destination'));
    destPanel.orientation = 'column';
    destPanel.alignChildren = 'left';
    destPanel.margins = [15, 20, 15, 10];

    var nameGroup = destPanel.add('group');
    nameGroup.orientation = 'row';
    nameGroup.alignChildren = ['left', 'center'];
    var nameLabel = nameGroup.add('statictext', undefined, L('layerName'));
    var etLayerName = nameGroup.add('edittext', undefined, '_mergedLayer');
    etLayerName.characters = 19;

    var colorGroup = destPanel.add('group');
    colorGroup.orientation = 'row';
    colorGroup.alignChildren = ['left', 'center'];
    var colorLabel = colorGroup.add('statictext', undefined, L('layerColor'));
    var colorSwatch = colorGroup.add('panel');
    colorSwatch.preferredSize = [14, 14];
    colorSwatch.minimumSize = [14, 14];
    var etLayerColor = colorGroup.add('edittext', undefined, '79,127,255');
    etLayerColor.characters = 12;

    function updateColorSwatch() {
        var colorValues = parseLayerColorValue(etLayerColor.text);
        colorSwatch.onDraw = function () {
            var g = this.graphics;
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255, 1]);
            var pen = g.newPen(g.PenType.SOLID_COLOR, [0.4, 0.4, 0.4, 1], 1);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.fillPath(brush);
            g.strokePath(pen);
        };
        if (colorSwatch.parent && colorSwatch.parent.layout) {
            colorSwatch.parent.layout.layout(true);
        }
    }

    etLayerColor.onChanging = updateColorSwatch;
    etLayerColor.onChange = updateColorSwatch;
    updateColorSwatch();

    var excludePanel = dlg.add('panel', undefined, L('exclude'));
    excludePanel.orientation = 'column';
    excludePanel.alignChildren = 'fill';
    excludePanel.margins = [15, 20, 15, 10];

    var excludeGroup = excludePanel.add('group');
    excludeGroup.orientation = 'row';
    excludeGroup.alignChildren = ['left','top'];
    excludeGroup.spacing = 15;

    // 左カラム：ロック
    var lockedPanel = excludeGroup.add('panel', undefined, 'ロック');
    lockedPanel.orientation = 'column';
    lockedPanel.alignChildren = 'left';
    lockedPanel.margins = [15, 20, 15, 10];

    var cbSkipLocked = lockedPanel.add('checkbox', undefined, L('skipLockedLayers'));
    cbSkipLocked.value = true;
    var cbSkipLockedObjects = lockedPanel.add('checkbox', undefined, L('skipLockedObjects'));
    cbSkipLockedObjects.value = true;

    // 右カラム：非表示
    var hiddenPanel = excludeGroup.add('panel', undefined, '非表示');
    hiddenPanel.orientation = 'column';
    hiddenPanel.alignChildren = 'left';
    hiddenPanel.margins = [15, 20, 15, 10];

    var cbSkipHidden = hiddenPanel.add('checkbox', undefined, L('skipHiddenLayers'));
    cbSkipHidden.value = true;
    var cbSkipHiddenObjects = hiddenPanel.add('checkbox', undefined, L('skipHiddenObjects'));
    cbSkipHiddenObjects.value = true;

    var optionsPanel = dlg.add('panel', undefined, L('options'));
    optionsPanel.orientation = 'column';
    optionsPanel.alignChildren = 'left';
    optionsPanel.margins = [15, 20, 15, 10];

    var cbDeleteEmpty = optionsPanel.add('checkbox', undefined, L('deleteEmptyLayers'));
    cbDeleteEmpty.value = true;

    var buttonGroup = dlg.add('group');
    buttonGroup.orientation = 'row';
    buttonGroup.alignment = ['center', 'center'];
    buttonGroup.alignChildren = ['center', 'center'];
    var btnCancel = buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var btnOk = buttonGroup.add('button', undefined, L('ok'), { name: 'ok' });

    if (dlg.show() !== 1) {
        return null;
    }

    return {
        skipLockedLayers: cbSkipLocked.value,
        skipHiddenLayers: cbSkipHidden.value,
        skipLockedObjects: cbSkipLockedObjects.value,
        skipHiddenObjects: cbSkipHiddenObjects.value,
        mergedLayerName: etLayerName.text,
        deleteEmptyLayers: cbDeleteEmpty.value,
        reuseExistingMergedLayer: true,
        layerColorValue: etLayerColor.text
    };
}

function getOrCreateMergedLayer(documentRef, mergedLayerName, reuseExistingMergedLayer) {
    var mergedLayer = null;

    if (reuseExistingMergedLayer) {
        try {
            mergedLayer = documentRef.layers.getByName(mergedLayerName);
        } catch (e) {
            mergedLayer = null;
        }
    }

    if (mergedLayer === null) {
        mergedLayer = documentRef.layers.add();
        mergedLayer.name = mergedLayerName;
    }

    return mergedLayer;
}

function clampColorValue(value) {
    var num = parseInt(value, 10);
    if (isNaN(num)) return 0;
    if (num < 0) return 0;
    if (num > 255) return 255;
    return num;
}

function parseLayerColorValue(layerColorValue) {
    var defaultColor = [79, 127, 255];
    if (!layerColorValue || layerColorValue === '') {
        return defaultColor;
    }

    var parts = String(layerColorValue).split(',');
    if (parts.length !== 3) {
        return defaultColor;
    }

    return [
        clampColorValue(parts[0]),
        clampColorValue(parts[1]),
        clampColorValue(parts[2])
    ];
}

function applyMergedLayerSettings(mergedLayer, layerColorValue) {
    mergedLayer.visible = true;
    mergedLayer.locked = false;
    mergedLayer.printable = true;
    var colorValues = parseLayerColorValue(layerColorValue);
    var customRGBColor = new RGBColor();
    customRGBColor.red = colorValues[0];
    customRGBColor.green = colorValues[1];
    customRGBColor.blue = colorValues[2];
    mergedLayer.color = customRGBColor;
}

function main() {
    var documentRef;
    try {
        documentRef = app.activeDocument;
    } catch (error) {
        return;
    }

    var allTopLayers = documentRef.layers;
    var mergedLayerName = "_mergedLayer";

    var hasExistingMergedLayer = false;

    try {
        documentRef.layers.getByName(mergedLayerName);
        hasExistingMergedLayer = true;
    } catch (e) {
        hasExistingMergedLayer = false;
    }

    var options = showOptionsDialog(hasExistingMergedLayer);
    if (!options) {
        return;
    }

    if (options.mergedLayerName && options.mergedLayerName !== '') {
        mergedLayerName = options.mergedLayerName;
    }

    var mergedLayer = getOrCreateMergedLayer(documentRef, mergedLayerName, options.reuseExistingMergedLayer);
    applyMergedLayerSettings(mergedLayer, options.layerColorValue);

    var topCount = allTopLayers.length;
    for (var i = topCount - 1; i >= 0; i--) {
        var currentLayer = allTopLayers[i];

        if (isExcludedLayer(currentLayer.name) || currentLayer === mergedLayer) {
            continue;
        }

        moveItemsToTargetLayer(currentLayer, mergedLayer, options);
    }

    if (options.deleteEmptyLayers) {
        var emptyLayerList = [];
        findEmptyLayers(documentRef, emptyLayerList, mergedLayer, options);
        deleteLayers(emptyLayerList);
    }
}

/*
 * 指定レイヤー配下の全ページアイテムを再帰的に destinationLayer へ移動。
 * Move all page items under sourceLayer recursively to destinationLayer.
 * Returns true if any items were moved.
 */
function moveItemsToTargetLayer(sourceLayer, destinationLayer, options) {
    if (sourceLayer === destinationLayer) return false;
    if (options.skipLockedLayers && sourceLayer.locked) return false;
    if (options.skipHiddenLayers && !sourceLayer.visible) return false;
    var didMove = false;

    var sublayers = sourceLayer.layers;
    for (var i = sublayers.length - 1; i >= 0; i--) {
        if (moveItemsToTargetLayer(sublayers[i], destinationLayer, options)) {
            didMove = true;
        }
    }

    var pageItems = sourceLayer.pageItems;
    // ロック項目は一時解除→移動→元に戻す。逆順走査 + PLACEATBEGINNING で前後関係を保ちやすくする / Temporarily unlock, move, then restore lock. Reverse traversal + PLACEATBEGINNING helps preserve stacking order.
    for (var j = pageItems.length - 1; j >= 0; j--) {
        var item = pageItems[j];
        var wasLocked = false;
        try {
            wasLocked = item.locked;
            if (options.skipLockedObjects && wasLocked) {
                continue;
            }
            if (options.skipHiddenObjects && item.hidden) {
                continue;
            }
            if (wasLocked) item.locked = false;
            item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
            didMove = true;
        } catch (error) {
            // 移動できない項目は無視 / ignore items that still cannot be moved
        } finally {
            try {
                if (wasLocked) item.locked = true;
            } catch (e2) { }
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
function findEmptyLayers(container, resultArray, mergedLayer, options) {
    var layers = container.layers;
    for (var i = 0; i < layers.length; i++) {
        var currentLayer = layers[i];

        if (currentLayer === mergedLayer) {
            continue;
        }
        if (options.skipLockedLayers && currentLayer.locked) {
            continue;
        }
        if (options.skipHiddenLayers && !currentLayer.visible) {
            continue;
        }

        if (currentLayer.layers.length > 0) {
            findEmptyLayers(currentLayer, resultArray, mergedLayer, options);
        }

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
        try { layerArray[i].remove(); } catch (e) { }
    }
}

main();