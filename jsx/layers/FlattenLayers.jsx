#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// スクリプトバージョン

var SCRIPT_VERSION = "v1.6.1";

/*
レイヤー統合（フラット化）を行うIllustrator用スクリプト。
除外名（例：bg/背景/background）を持つレイヤーを残しつつ、その他のレイヤー／サブレイヤー配下の全オブジェクトを
「_mergedLayer」に集約（移動）してフラット化します。その後、空になったレイヤーを再帰的に検出して削除します。

### スクリプト名：

レイヤー統合（フラット化） / Flatten Layers

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### 概要：

- 実行前にダイアログを表示し、処理条件を選択可能
- 除外レイヤーを残し、それ以外のレイヤー／サブレイヤー配下のオブジェクトを「_mergedLayer」に移動してフラット化
- ロック / 非表示のレイヤー・オブジェクトをそれぞれ対象外にするか選択可能
- 必要に応じて、中身が残ったサブレイヤーを最上位のレイヤーへ移動可能
- 空のレイヤー／サブレイヤーを、削除可能なものがなくなるまで再帰反復で削除
- まとめ先のレイヤー名、既存まとめ先の再利用、レイヤーカラーを指定可能

### 処理の流れ：

1. ドキュメント取得（未オープンなら終了）
2. ダイアログで処理条件を選択（キャンセル時は終了）
3. 設定に応じて既存「_mergedLayer」を取得、または新規作成
4. 除外レイヤーを除いて、条件に合う全レイヤー／サブレイヤー配下のオブジェクトを「_mergedLayer」に移動してフラット化
5. 必要に応じて、中身が残ったサブレイヤーを最上位のレイヤーへ移動
6. 設定が ON の場合、空のレイヤー／サブレイヤーがなくなるまで再帰反復で削除

### 更新履歴：

- v1.0 (20250414) : 初期バージョン

Illustrator script to flatten layers. It consolidates all objects (except in excluded layers
such as bg/背景/background) from layers and sublayers into a single layer named "_mergedLayer", then recursively
removes layers that became empty.

### Script Name:

Flatten Layers

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### Overview:

- Show a dialog before execution so the processing conditions can be selected
- Flatten objects from non-excluded layers and sublayers into "_mergedLayer"
- Allow the user to choose whether locked / hidden layers and objects are excluded
- Optionally move remaining non-empty sublayers to the top level
- Delete empty layers / sublayers repeatedly until no more removable empty layers remain
- Let the user specify the destination layer name, reuse an existing destination layer, and set the destination layer color

### Process Flow:

1. Get the active document (exit if none)
2. Show the options dialog (exit if canceled)
3. Reuse or create "_mergedLayer" according to the selected options
4. Flatten items from eligible non-excluded layers and sublayers into "_mergedLayer"
5. Optionally move remaining non-empty sublayers to the top level
6. If enabled, repeatedly delete empty layers / sublayers until none remain

### Update History:

- v1.0 (20250414) : Initial release

*/

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
    process: {
        ja: '処理',
        en: 'Process'
    },
    promoteSublayers: {
        ja: '中身が残るサブレイヤーを最上位化',
        en: 'Promote remaining sublayers to the top level'
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
    reuseExistingMergedLayer: {
        ja: '既存のまとめ先を再利用',
        en: 'Reuse existing destination layer'
    },
    lockedPanelTitle: {
        ja: 'ロック',
        en: 'Locked'
    },
    hiddenPanelTitle: {
        ja: '非表示',
        en: 'Hidden'
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
    toggleAllExclusions: {
        ja: '対象外設定を一括切替',
        en: 'Toggle all on/off'
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

function createProcessStats() {
    return {
        moveFailureCount: 0,
        deleteFailureCount: 0,
        visibilityRestoreFailureCount: 0,
        layerLockRestoreFailureCount: 0,
        itemLockRestoreFailureCount: 0,
        topLevelPromotionFailureCount: 0,
        parentAccessFailureCount: 0
    };
}

function showOptionsDialog(documentRef, hasExistingMergedLayer) {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';

    var processPanel = dlg.add('panel', undefined, L('process'));
    processPanel.orientation = 'column';
    processPanel.alignChildren = 'fill';
    processPanel.margins = [15, 20, 15, 10];

    var cbPromoteSublayers = processPanel.add('checkbox', undefined, L('promoteSublayers'));
    cbPromoteSublayers.value = true;

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

    var cbReuseExistingMergedLayer = destPanel.add('checkbox', undefined, L('reuseExistingMergedLayer'));
    cbReuseExistingMergedLayer.value = hasExistingMergedLayer;
    cbReuseExistingMergedLayer.enabled = hasExistingMergedLayer;

    function updateReuseExistingMergedLayerState() {
        var layerName = etLayerName.text;
        var hasMatchingLayer = false;

        if (layerName && layerName !== '') {
            try {
                documentRef.layers.getByName(layerName);
                hasMatchingLayer = true;
            } catch (e) {
                hasMatchingLayer = false;
            }
        }

        cbReuseExistingMergedLayer.enabled = hasMatchingLayer;
        if (!hasMatchingLayer) {
            cbReuseExistingMergedLayer.value = false;
        }
    }

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

    etLayerName.onChanging = updateReuseExistingMergedLayerState;
    etLayerName.onChange = updateReuseExistingMergedLayerState;
    updateReuseExistingMergedLayerState();

    var excludePanel = processPanel.add('panel', undefined, L('exclude'));
    excludePanel.orientation = 'column';
    excludePanel.alignChildren = 'fill';
    excludePanel.margins = [15, 20, 15, 10];

    var excludeGroup = excludePanel.add('group');
    excludeGroup.orientation = 'row';
    excludeGroup.alignChildren = ['left', 'top'];
    excludeGroup.spacing = 15;

    // 左カラム：ロック
    var lockedPanel = excludeGroup.add('panel', undefined, L('lockedPanelTitle'));
    lockedPanel.orientation = 'column';
    lockedPanel.alignChildren = 'left';
    lockedPanel.margins = [15, 20, 15, 10];

    var cbSkipLocked = lockedPanel.add('checkbox', undefined, L('skipLockedLayers'));
    cbSkipLocked.value = true;
    var cbSkipLockedObjects = lockedPanel.add('checkbox', undefined, L('skipLockedObjects'));
    cbSkipLockedObjects.value = true;

    // 右カラム：非表示
    var hiddenPanel = excludeGroup.add('panel', undefined, L('hiddenPanelTitle'));
    hiddenPanel.orientation = 'column';
    hiddenPanel.alignChildren = 'left';
    hiddenPanel.margins = [15, 20, 15, 10];

    var cbSkipHidden = hiddenPanel.add('checkbox', undefined, L('skipHiddenLayers'));
    cbSkipHidden.value = true;
    var cbSkipHiddenObjects = hiddenPanel.add('checkbox', undefined, L('skipHiddenObjects'));
    cbSkipHiddenObjects.value = true;

    var toggleAllGroup = excludePanel.add('group');
    toggleAllGroup.orientation = 'row';
    toggleAllGroup.alignment = ['left', 'top'];
    var cbToggleAllExclusions = toggleAllGroup.add('checkbox', undefined, L('toggleAllExclusions'));
    cbToggleAllExclusions.value = true;

    function updateToggleAllExclusionsState() {
        cbToggleAllExclusions.value = cbSkipLocked.value && cbSkipLockedObjects.value && cbSkipHidden.value && cbSkipHiddenObjects.value;
    }

    cbToggleAllExclusions.onClick = function () {
        var newValue = cbToggleAllExclusions.value;
        cbSkipLocked.value = newValue;
        cbSkipLockedObjects.value = newValue;
        cbSkipHidden.value = newValue;
        cbSkipHiddenObjects.value = newValue;
    };

    cbSkipLocked.onClick = updateToggleAllExclusionsState;
    cbSkipLockedObjects.onClick = updateToggleAllExclusionsState;
    cbSkipHidden.onClick = updateToggleAllExclusionsState;
    cbSkipHiddenObjects.onClick = updateToggleAllExclusionsState;

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
    buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    buttonGroup.add('button', undefined, L('ok'), { name: 'ok' });

    if (dlg.show() !== 1) {
        return null;
    }

    return {
        skipLockedLayers: cbSkipLocked.value,
        skipHiddenLayers: cbSkipHidden.value,
        skipLockedObjects: cbSkipLockedObjects.value,
        skipHiddenObjects: cbSkipHiddenObjects.value,
        promoteSublayersToTopLevel: cbPromoteSublayers.value,
        mergedLayerName: etLayerName.text,
        deleteEmptyLayers: cbDeleteEmpty.value,
        reuseExistingMergedLayer: cbReuseExistingMergedLayer.value,
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

    var mergedLayerName = "_mergedLayer";
    var hasExistingMergedLayer = false;

    try {
        documentRef.layers.getByName(mergedLayerName);
        hasExistingMergedLayer = true;
    } catch (e) {
        hasExistingMergedLayer = false;
    }

    var options = showOptionsDialog(documentRef, hasExistingMergedLayer);
    if (!options) {
        return;
    }

    if (options.mergedLayerName && options.mergedLayerName !== '') {
        mergedLayerName = options.mergedLayerName;
    }

    var stats = createProcessStats();
    var mergedLayer = getOrCreateMergedLayer(documentRef, mergedLayerName, options.reuseExistingMergedLayer);
    applyMergedLayerSettings(mergedLayer, options.layerColorValue);

    var allTopLayers = documentRef.layers;
    for (var i = allTopLayers.length - 1; i >= 0; i--) {
        var currentLayer = allTopLayers[i];
        if (currentLayer === mergedLayer) {
            continue;
        }
        if (isExcludedLayer(currentLayer.name)) {
            continue;
        }
        moveItemsToTargetLayer(currentLayer, mergedLayer, options, stats);
    }

    if (options.promoteSublayersToTopLevel) {
        promoteSublayersToTop(documentRef, mergedLayer, options, stats);
    }

    if (options.deleteEmptyLayers) {
        deleteEmptyLayersRecursively(documentRef, mergedLayer, options, stats);
    }

    var messages = [];
    if (stats.moveFailureCount > 0) {
        messages.push((lang === 'ja' ? '移動失敗' : 'Move failures') + ': ' + stats.moveFailureCount);
    }
    if (stats.deleteFailureCount > 0) {
        messages.push((lang === 'ja' ? '削除失敗' : 'Delete failures') + ': ' + stats.deleteFailureCount);
    }
    if (stats.visibilityRestoreFailureCount > 0) {
        messages.push((lang === 'ja' ? '表示状態の復元失敗' : 'Visibility restore failures') + ': ' + stats.visibilityRestoreFailureCount);
    }
    if (stats.layerLockRestoreFailureCount > 0) {
        messages.push((lang === 'ja' ? 'レイヤーロック復元失敗' : 'Layer lock restore failures') + ': ' + stats.layerLockRestoreFailureCount);
    }
    if (stats.itemLockRestoreFailureCount > 0) {
        messages.push((lang === 'ja' ? 'オブジェクトロック復元失敗' : 'Item lock restore failures') + ': ' + stats.itemLockRestoreFailureCount);
    }
    if (stats.topLevelPromotionFailureCount > 0) {
        messages.push((lang === 'ja' ? '最上位化失敗' : 'Top-level promotion failures') + ': ' + stats.topLevelPromotionFailureCount);
    }
    if (stats.parentAccessFailureCount > 0) {
        messages.push((lang === 'ja' ? '親レイヤーアクセス失敗' : 'Parent access failures') + ': ' + stats.parentAccessFailureCount);
    }
    if (messages.length > 0) {
        alert(messages.join('\n'));
    }
}


/*
 * 指定レイヤー配下を再帰走査し、各レイヤー直下のページアイテムだけを destinationLayer へ移動。
 * Recursively walk sourceLayer and move only the page items directly under each layer to destinationLayer.
 */
function moveItemsToTargetLayer(sourceLayer, destinationLayer, options, stats) {
    if (sourceLayer === destinationLayer) return false;
    if (options.skipLockedLayers && sourceLayer.locked) return false;
    if (options.skipHiddenLayers && !sourceLayer.visible) return false;

    var didMove = false;

    var restoreVisibility = false;
    var originalVisibility = true;

    var restoreLayerLock = false;
    var originalLayerLock = false;

    if (!options.skipHiddenLayers && !sourceLayer.visible) {
        try {
            originalVisibility = sourceLayer.visible;
            sourceLayer.visible = true;
            restoreVisibility = true;
        } catch (e0) { }
    }

    if (!options.skipLockedLayers && sourceLayer.locked) {
        try {
            originalLayerLock = sourceLayer.locked;
            sourceLayer.locked = false;
            restoreLayerLock = true;
        } catch (e00) { }
    }

    for (var i = sourceLayer.layers.length - 1; i >= 0; i--) {
        if (moveItemsToTargetLayer(sourceLayer.layers[i], destinationLayer, options, stats)) {
            didMove = true;
        }
    }

    var pageItems = sourceLayer.pageItems;
    for (var j = pageItems.length - 1; j >= 0; j--) {
        var item = pageItems[j];
        if (item.parent !== sourceLayer) {
            continue;
        }

        var wasLocked = false;
        try {
            wasLocked = item.locked;
            if (options.skipLockedObjects && wasLocked) {
                continue;
            }
            if (shouldSkipHiddenItem(item, options)) {
                continue;
            }
            if (wasLocked) item.locked = false;
            item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
            didMove = true;
        } catch (error) {
            if (stats) {
                stats.moveFailureCount++;
            }
            // 移動できない項目は無視 / ignore items that still cannot be moved
        } finally {
            try {
                if (wasLocked) item.locked = true;
            } catch (e2) {
                if (stats) {
                    stats.itemLockRestoreFailureCount++;
                }
            }
        }
    }

    if (restoreVisibility) {
        try {
            sourceLayer.visible = originalVisibility;
        } catch (e3) {
            if (stats) {
                stats.visibilityRestoreFailureCount++;
            }
        }
    }
    if (restoreLayerLock) {
        try {
            sourceLayer.locked = originalLayerLock;
        } catch (e4) {
            if (stats) {
                stats.layerLockRestoreFailureCount++;
            }
        }
    }

    return didMove;
}

function shouldSkipHiddenItem(item, options) {
    if (!options.skipHiddenObjects) {
        return false;
    }
    if (!item.hidden) {
        return false;
    }
    if (!options.skipHiddenLayers && isHiddenOnlyByLayerVisibility(item)) {
        return false;
    }
    return true;
}

function isHiddenOnlyByLayerVisibility(item) {
    var current = item;
    while (current && current.parent) {
        current = current.parent;
        if (current.typename === 'Layer' && !current.visible) {
            return true;
        }
    }
    return false;
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
// 中身が残るサブレイヤーをトップレベルへ移動
// ==========================
function promoteSublayersToTop(documentRef, mergedLayer, options, stats) {
    var guard = 0;
    while (guard < 1000) {
        var batch = [];
        collectDirectSublayers(documentRef, batch, mergedLayer, options);
        if (batch.length === 0) {
            break;
        }

        for (var i = 0; i < batch.length; i++) {
            moveLayerToTopLevel(documentRef, batch[i], mergedLayer, stats);
        }
        guard++;
    }
}

function collectDirectSublayers(container, resultArray, mergedLayer, options) {
    var layers = container.layers;

    for (var i = layers.length - 1; i >= 0; i--) {
        var currentLayer = layers[i];

        if (currentLayer === mergedLayer) {
            continue;
        }
        if (isExcludedLayer(currentLayer.name)) {
            continue;
        }
        if (options.skipLockedLayers && currentLayer.locked) {
            continue;
        }
        if (options.skipHiddenLayers && !currentLayer.visible && !(currentLayer.parent && currentLayer.parent.typename === 'Layer')) {
            continue;
        }

        if (currentLayer.parent && currentLayer.parent.typename === 'Layer' && layerHasRemainingContent(currentLayer)) {
            resultArray.push(currentLayer);
        }

        if (currentLayer.layers.length > 0) {
            collectDirectSublayers(currentLayer, resultArray, mergedLayer, options);
        }
    }
}

function layerHasRemainingContent(layerRef) {
    return hasDirectPageItems(layerRef) || layerRef.layers.length > 0;
}

function moveLayerToTopLevel(documentRef, layerRef, mergedLayer, stats) {
    if (!layerRef || !layerRef.parent || layerRef.parent.typename !== 'Layer') {
        return;
    }

    var anchor = getTopLevelMoveAnchor(documentRef, mergedLayer, layerRef);
    if (!anchor) {
        return;
    }

    try {
        layerRef.move(anchor, ElementPlacement.PLACEAFTER);
    } catch (e) {
        if (stats) {
            stats.topLevelPromotionFailureCount++;
        }
    }
}

function getTopLevelMoveAnchor(documentRef, mergedLayer, movingLayer) {
    var topLayers = documentRef.layers;
    for (var i = topLayers.length - 1; i >= 0; i--) {
        var candidate = topLayers[i];
        if (candidate === movingLayer) {
            continue;
        }
        if (candidate !== mergedLayer) {
            return candidate;
        }
    }
    return mergedLayer || null;
}

// ==========================
// 空レイヤー削除の制御関数 / Control loop for empty-layer deletion
// - findEmptyLayers() : 現時点で空のレイヤー／サブレイヤーを収集
// - deleteLayers()    : 収集済みレイヤーを実際に削除
// - この関数          : 親レイヤーが後から空になるケースに備え、空がなくなるまで反復
// ==========================
function deleteEmptyLayersRecursively(documentRef, mergedLayer, options, stats) {
    var guard = 0;
    while (guard < 1000) {
        var emptyLayerList = [];
        findEmptyLayers(documentRef, emptyLayerList, mergedLayer, options);
        if (emptyLayerList.length === 0) {
            break;
        }
        deleteLayers(emptyLayerList, stats);
        guard++;
    }
}

// ==========================
// 空レイヤー収集関数 / Collect empty layers recursively
// - 削除は行わず、削除候補の収集だけを担当
// - mergedLayer と除外名レイヤーは対象外
// - トップレベルの非表示 / ロック除外はここで判定
// - 子を先に走査し、現時点で空のレイヤーを resultArray に積む
// ==========================
function findEmptyLayers(container, resultArray, mergedLayer, options) {
    var layers = container.layers;
    for (var i = 0; i < layers.length; i++) {
        var currentLayer = layers[i];

        if (currentLayer === mergedLayer) {
            continue;
        }
        if (isExcludedLayer(currentLayer.name)) {
            continue;
        }

        var isTopLevelLayer = !(currentLayer.parent && currentLayer.parent.typename === 'Layer');
        var skipCurrentLayerDeletion = false;

        if (options.skipLockedLayers && currentLayer.locked && isTopLevelLayer) {
            skipCurrentLayerDeletion = true;
        }
        if (options.skipHiddenLayers && !currentLayer.visible && isTopLevelLayer) {
            skipCurrentLayerDeletion = true;
        }

        if (currentLayer.layers.length > 0) {
            findEmptyLayers(currentLayer, resultArray, mergedLayer, options);
        }

        if (skipCurrentLayerDeletion) {
            continue;
        }

        var hasItems = hasDirectPageItems(currentLayer);
        var hasChildLayers = currentLayer.layers.length > 0;
        if (!hasItems && !hasChildLayers) {
            resultArray.push(currentLayer);
        }
    }
}

// 直属アイテム判定専用 / Check only whether the layer itself directly owns pageItems
// 子サブレイヤー配下のアイテムは含めない。
function hasDirectPageItems(layerRef) {
    var items = layerRef.pageItems;
    for (var i = 0; i < items.length; i++) {
        if (items[i].parent === layerRef) {
            return true;
        }
    }
    return false;
}

// ==========================
// 空レイヤー削除実行関数 / Delete listed layers
// - 収集済みの削除候補だけを削除
// - 深いサブレイヤーから先に削除
// - 削除前に親レイヤーを一時的に visible / unlocked にして到達可能にする
// - 収集ロジックや再試行制御は持たない
// ==========================
function deleteLayers(layerArray, stats) {
    var sortedLayers = layerArray.slice(0);
    sortedLayers.sort(function (a, b) {
        return getLayerDepth(b) - getLayerDepth(a);
    });

    for (var i = 0; i < sortedLayers.length; i++) {
        var layerRef = sortedLayers[i];
        try {
            if (!ensureLayerParentsAccessible(layerRef, stats)) {
                if (stats) {
                    stats.parentAccessFailureCount++;
                }
                continue;
            }
            layerRef.visible = true;
            layerRef.locked = false;
            layerRef.remove();
        } catch (e) {
            if (stats) {
                stats.deleteFailureCount++;
            }
        }
    }
}

function getLayerDepth(layerRef) {
    var depth = 0;
    var current = layerRef;
    while (current && current.parent && current.parent.typename === 'Layer') {
        depth++;
        current = current.parent;
    }
    return depth;
}

function ensureLayerParentsAccessible(layerRef, stats) {
    var chain = [];
    var current = layerRef;
    while (current && current.parent && current.parent.typename === 'Layer') {
        current = current.parent;
        chain.unshift(current);
    }

    for (var i = 0; i < chain.length; i++) {
        try {
            chain[i].visible = true;
            chain[i].locked = false;
        } catch (e) {
            return false;
        }
    }
    return true;
}

main();