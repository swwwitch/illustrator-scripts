#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// スクリプトバージョン

var SCRIPT_VERSION = "v1.7.3";

/*
レイヤー統合（フラット化）を行うIllustrator用スクリプト。

- 除外名（bg）を持つレイヤーを残しつつ、その他のレイヤー／サブレイヤー配下のオブジェクトを指定したまとめ先レイヤーへ集約（移動）してフラット化します。必要に応じて、中身が残ったサブレイヤーを上位レベルのレイヤーへ移動し、空になったレイヤーを再帰的に検出して削除します。
- ロック / 非表示のレイヤー・オブジェクトは個別に対象外指定でき、該当するものが存在しない項目はダイアログ上で自動的にディム表示されます。
- ガイドは「統合」「現在のレイヤーに保持」「別レイヤーに移動」の3つのモードから選択できます。

### スクリプト名：

レイヤー統合（フラット化） / Flatten Layers

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### 作成日：

2025-04-14

### 更新日：

2026-04-13

### 概要：

- 実行前にダイアログを表示し、処理条件を選択可能
- 除外レイヤーを残し、それ以外のレイヤー／サブレイヤー配下のオブジェクトを指定レイヤーへ移動してフラット化
- ロック / 非表示のレイヤー・オブジェクトをそれぞれ対象外にするか選択可能
- ロックレイヤー / 非表示レイヤー / 非表示オブジェクト / ガイド / サブレイヤーが存在しない場合、対応するUIを自動的にディム表示
- 対象外設定の一括切替は、有効な項目だけを一括で ON / OFF
- 必要に応じて、中身が残ったサブレイヤーを上位レベルのレイヤーへ移動可能
- ガイドの扱いを「統合」「現在のレイヤーに保持」「別レイヤーに移動」から選択可能
- 「現在のレイヤーに保持」では、サブレイヤー直下のガイドを1つ上のレイヤーへ繰り上げて保持
- 「別レイヤーに移動」では、統合後のガイドを指定レイヤーへ移動
- 空のレイヤー／サブレイヤーを、削除可能なものがなくなるまで再帰反復で削除
- skipLockedLayers / skipHiddenLayers はトップレベルレイヤーとサブレイヤーの両方に一貫適用
- ガイドだけ残っているレイヤーは空レイヤーとは見なさない
- まとめ先のレイヤー名、既存まとめ先の再利用、レイヤーカラーを指定可能
- 既存まとめ先を再利用しない場合、同名レイヤーが存在すれば連番付きの別名で新規作成

### 処理の流れ：

1. ドキュメント取得（未オープンなら終了）
2. ダイアログで処理条件を選択（該当しない項目は自動的にディム表示、キャンセル時は終了）
3. 設定に応じて既存まとめ先レイヤーを取得、または一意な名前で新規作成
4. 除外レイヤーを除いて、条件に合う全レイヤー／サブレイヤー配下のオブジェクトをまとめ先レイヤーへ移動してフラット化
5. ガイドモードが「現在のレイヤーに保持」の場合、サブレイヤー直下のガイドを上位レイヤーへ繰り上げ
6. 必要に応じて、中身が残ったサブレイヤーを上位レベルのレイヤーへ移動
7. ガイドモードが「別レイヤーに移動」の場合、統合後のガイドを指定レイヤーへ移動
8. 設定が ON の場合、空のレイヤー／サブレイヤーがなくなるまで再帰反復で削除
9. 失敗件数があれば、処理後に件数だけを簡潔に通知

### 更新履歴：

- v1.0 (20250414) : 初期バージョン
- v1.7.1 (20260408) : 対象外UIの自動ディム表示、一括切替の有効項目限定、skipLockedLayers / skipHiddenLayers のサブレイヤーまでの一貫適用、ガイド panel 初期化の明確化、概要とコメントの更新
- v1.7.3 (20260413) : 除外レイヤー名を「bg」のみに変更（「背景」「background」を削除）

Illustrator script to flatten layers. It keeps excluded layers (bg),
moves objects under all other layers and sublayers into a specified destination layer, optionally
promotes remaining non-empty sublayers to the top level, and recursively deletes layers that become empty.
Locked / hidden layers and objects can be excluded independently, and dialog items with no matching targets
are dimmed automatically. Guides can be handled in one of three modes: integrate, keep in the current layer,
or move to another layer.

### Script Name:

Flatten Layers

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/FlattenLayers.jsx

### Created:

2025-04-14

### Updated:

2026-04-13

### Overview:

- Show a dialog before execution so the processing conditions can be selected
- Flatten objects from non-excluded layers and sublayers into the specified destination layer
- Let the user choose whether locked / hidden layers and objects are excluded
- Automatically dim UI items when there are no locked layers, hidden layers, hidden objects, guides, or sublayers in the document
- Toggle all exclusion options on / off only for the items that are currently enabled
- Optionally move remaining non-empty sublayers to the top level
- Let the user choose how guides are handled: integrate, keep in the current layer, or move to another layer
- In “Keep in the current layer” mode, guides directly under sublayers are hoisted to the parent layer
- In “Move to another layer” mode, guides are moved to the specified guide layer after flattening
- Delete empty layers / sublayers repeatedly until no more removable empty layers remain
- Apply skipLockedLayers / skipHiddenLayers consistently to both top-level layers and sublayers
- Layers that still contain only guides are not treated as empty
- Let the user specify the destination layer name, reuse an existing destination layer, and set the destination layer color
- If destination-layer reuse is off and the same name already exists, create a new layer with a numbered unique name

### Process Flow:

1. Get the active document (exit if none)
2. Show the options dialog, automatically dim irrelevant items, and exit if canceled
3. Reuse the existing destination layer or create a new uniquely named one according to the selected options
4. Flatten items from eligible non-excluded layers and sublayers into the destination layer
5. If the guide mode is “Keep in the current layer,” hoist guides directly under sublayers to the parent layer
6. Optionally move remaining non-empty sublayers to the top level
7. If the guide mode is “Move to another layer,” move guides from the flattened result into the specified guide layer
8. If enabled, repeatedly delete empty layers / sublayers until none remain
9. If any failures occurred, show only the failure counts after processing

### Update History:

- v1.0 (20250414) : Initial release
- v1.7.1 (20260408) : Added automatic dimming of irrelevant exclusion UI, enabled-only toggle-all behavior, consistent skipLockedLayers / skipHiddenLayers handling for sublayers, clearer guides-panel initialization, and updated overview/comments
- v1.7.3 (20260413) : Narrowed excluded layer names to only "bg" (removed "背景" and "background")

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
        ja: 'サブレイヤーを上位レベルのレイヤーに変更',
        en: 'Move sublayers to top-level layers'
    },
    guides: {
        ja: 'ガイド',
        en: 'Guides'
    },
    separateGuides: {
        ja: '別レイヤーに移動',
        en: 'Move to another layer'
    },
    keepGuidesInCurrentLayer: {
        ja: '現在のレイヤーに保持',
        en: 'Keep in the current layer'
    },
    integrateGuides: {
        ja: '統合',
        en: 'Integrate'
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
    layersPanelTitle: {
        ja: 'レイヤー',
        en: 'Layers'
    },
    hiddenPanelTitle: {
        ja: '非表示',
        en: 'Hidden'
    },
    slashSlashLayer: {
        ja: '「//」レイヤー',
        en: '“//” layers'
    },
    objectsPanelTitle: {
        ja: 'オブジェクト',
        en: 'Objects'
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
        en: 'Toggle all exclusions'
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
        guideSeparationFailureCount: 0,
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

    function documentHasAnyLockedLayers(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layerHasAnyLockedLayers(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function layerHasAnyLockedLayers(layerRef) {
        if (layerRef.locked) {
            return true;
        }

        var childLayers = layerRef.layers;
        for (var i = 0; i < childLayers.length; i++) {
            if (layerHasAnyLockedLayers(childLayers[i])) {
                return true;
            }
        }

        return false;
    }

    function documentHasAnyHiddenLayers(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layerHasAnyHiddenLayers(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function layerHasAnyHiddenLayers(layerRef) {
        if (!layerRef.visible) {
            return true;
        }

        var childLayers = layerRef.layers;
        for (var i = 0; i < childLayers.length; i++) {
            if (layerHasAnyHiddenLayers(childLayers[i])) {
                return true;
            }
        }

        return false;
    }

    function documentHasAnyHiddenObjects(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layerHasAnyHiddenObjects(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function documentHasAnySlashSlashLayers(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layerHasAnySlashSlashLayers(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function layerHasAnySlashSlashLayers(layerRef) {
        if (isSlashSlashLayer(layerRef)) {
            return true;
        }

        var childLayers = layerRef.layers;
        for (var i = 0; i < childLayers.length; i++) {
            if (layerHasAnySlashSlashLayers(childLayers[i])) {
                return true;
            }
        }

        return false;
    }

    function layerHasAnyHiddenObjects(layerRef) {
        var pageItems = layerRef.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].parent === layerRef && isActuallyHiddenItem(pageItems[i])) {
                return true;
            }
        }

        var childLayers = layerRef.layers;
        for (var j = 0; j < childLayers.length; j++) {
            if (layerHasAnyHiddenObjects(childLayers[j])) {
                return true;
            }
        }

        return false;
    }

    function isActuallyHiddenItem(item) {
        try {
            return item.hidden === true;
        } catch (e) {
            return false;
        }
    }

    function documentHasAnySublayers(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layers[i].parent && layers[i].parent.typename === 'Layer') {
                return true;
            }
            if (layers[i].layers && layers[i].layers.length > 0) {
                return true;
            }
            if (documentHasAnySublayers(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function documentHasAnyGuides(container) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            if (layerHasAnyGuides(layers[i])) {
                return true;
            }
        }
        return false;
    }

    function layerHasAnyGuides(layerRef) {
        var pageItems = layerRef.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].parent === layerRef && isGuideItem(pageItems[i])) {
                return true;
            }
        }

        var childLayers = layerRef.layers;
        for (var j = 0; j < childLayers.length; j++) {
            if (layerHasAnyGuides(childLayers[j])) {
                return true;
            }
        }

        return false;
    }

    var hasAnySublayers = documentHasAnySublayers(documentRef);
    var hasAnyLockedLayers = documentHasAnyLockedLayers(documentRef);
    var hasAnyHiddenLayers = documentHasAnyHiddenLayers(documentRef);
    var hasAnyHiddenObjects = documentHasAnyHiddenObjects(documentRef);
    var hasAnySlashSlashLayers = documentHasAnySlashSlashLayers(documentRef);
    var hasAnyGuides = documentHasAnyGuides(documentRef);
    cbPromoteSublayers.enabled = hasAnySublayers;
    if (!hasAnySublayers) cbPromoteSublayers.value = false;

    // ガイド設定パネル / Guides panel
    var guidesPanel = processPanel.add('panel', undefined, L('guides'));
    guidesPanel.orientation = 'column';
    guidesPanel.alignChildren = 'left';
    guidesPanel.margins = [15, 20, 15, 10];
    guidesPanel.enabled = hasAnyGuides;

    var rbIntegrateGuides = guidesPanel.add('radiobutton', undefined, L('integrateGuides'));
    rbIntegrateGuides.value = false;

    var rbKeepGuidesInCurrentLayer = guidesPanel.add('radiobutton', undefined, L('keepGuidesInCurrentLayer'));
    rbKeepGuidesInCurrentLayer.value = false;

    var separateGuidesGroup = guidesPanel.add('group');
    separateGuidesGroup.orientation = 'row';
    separateGuidesGroup.alignChildren = ['left', 'center'];

    var rbSeparateGuides = separateGuidesGroup.add('radiobutton', undefined, L('separateGuides'));
    rbSeparateGuides.value = true;

    var etGuideLayerName = separateGuidesGroup.add('edittext', undefined, '_guide');
    etGuideLayerName.characters = 12;

    function updateGuideOptionsState(selected) {
        if (selected === 'integrate' && rbIntegrateGuides.value) {
            rbKeepGuidesInCurrentLayer.value = false;
            rbSeparateGuides.value = false;
        } else if (selected === 'keep' && rbKeepGuidesInCurrentLayer.value) {
            rbIntegrateGuides.value = false;
            rbSeparateGuides.value = false;
        } else if (selected === 'separate' && rbSeparateGuides.value) {
            rbIntegrateGuides.value = false;
            rbKeepGuidesInCurrentLayer.value = false;
        }

        etGuideLayerName.enabled = rbSeparateGuides.value;
    }

    rbIntegrateGuides.onClick = function () {
        updateGuideOptionsState('integrate');
    };
    rbKeepGuidesInCurrentLayer.onClick = function () {
        updateGuideOptionsState('keep');
    };
    rbSeparateGuides.onClick = function () {
        updateGuideOptionsState('separate');
    };

    if (hasAnyGuides) {
        rbIntegrateGuides.value = false;
        rbKeepGuidesInCurrentLayer.value = false;
        rbSeparateGuides.value = true;
        updateGuideOptionsState('separate');
    } else {
        rbIntegrateGuides.value = true;
        rbKeepGuidesInCurrentLayer.value = false;
        rbSeparateGuides.value = false;
        updateGuideOptionsState('integrate');
    }

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
        var layerName = normalizeLayerName(etLayerName.text, '');
        var hasMatchingLayer = false;

        if (layerName !== '') {
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

    // 左カラム：レイヤー / Left column: layers
    var layerExcludePanel = excludeGroup.add('panel', undefined, L('layersPanelTitle'));
    layerExcludePanel.orientation = 'column';
    layerExcludePanel.alignChildren = 'left';
    layerExcludePanel.margins = [15, 20, 15, 10];

    var cbSkipLocked = layerExcludePanel.add('checkbox', undefined, L('lockedPanelTitle'));
    cbSkipLocked.value = false;
    cbSkipLocked.enabled = hasAnyLockedLayers;
    if (!hasAnyLockedLayers) cbSkipLocked.value = false;

    var cbSkipHidden = layerExcludePanel.add('checkbox', undefined, L('hiddenPanelTitle'));
    cbSkipHidden.value = false;
    cbSkipHidden.enabled = hasAnyHiddenLayers;
    if (!hasAnyHiddenLayers) cbSkipHidden.value = false;

    var cbSkipSlashSlashLayers = layerExcludePanel.add('checkbox', undefined, L('slashSlashLayer'));
    cbSkipSlashSlashLayers.value = false;
    cbSkipSlashSlashLayers.enabled = hasAnySlashSlashLayers;
    if (!hasAnySlashSlashLayers) cbSkipSlashSlashLayers.value = false;

    // 右カラム：オブジェクト / Right column: objects
    var objectExcludePanel = excludeGroup.add('panel', undefined, L('objectsPanelTitle'));
    objectExcludePanel.orientation = 'column';
    objectExcludePanel.alignChildren = 'left';
    objectExcludePanel.margins = [15, 20, 15, 10];

    var cbSkipLockedObjects = objectExcludePanel.add('checkbox', undefined, L('lockedPanelTitle'));
    cbSkipLockedObjects.value = false;

    var cbSkipHiddenObjects = objectExcludePanel.add('checkbox', undefined, L('hiddenPanelTitle'));
    cbSkipHiddenObjects.value = false;
    cbSkipHiddenObjects.enabled = hasAnyHiddenObjects;
    if (!hasAnyHiddenObjects) cbSkipHiddenObjects.value = false;

    var toggleAllGroup = excludePanel.add('group');
    toggleAllGroup.orientation = 'row';
    toggleAllGroup.alignment = ['left', 'top'];
    var cbToggleAllExclusions = toggleAllGroup.add('checkbox', undefined, L('toggleAllExclusions'));
    cbToggleAllExclusions.value = false;

    function updateToggleAllExclusionsState() {
        var enabledValues = [];

        if (cbSkipLocked.enabled) {
            enabledValues.push(cbSkipLocked.value);
        }
        if (cbSkipHidden.enabled) {
            enabledValues.push(cbSkipHidden.value);
        }
        if (cbSkipSlashSlashLayers.enabled) {
            enabledValues.push(cbSkipSlashSlashLayers.value);
        }
        if (cbSkipLockedObjects.enabled) {
            enabledValues.push(cbSkipLockedObjects.value);
        }
        if (cbSkipHiddenObjects.enabled) {
            enabledValues.push(cbSkipHiddenObjects.value);
        }

        cbToggleAllExclusions.value = enabledValues.length > 0 && allValuesAreTrue(enabledValues);
    }

    function allValuesAreTrue(values) {
        for (var i = 0; i < values.length; i++) {
            if (!values[i]) {
                return false;
            }
        }
        return true;
    }

    cbToggleAllExclusions.onClick = function () {
        var newValue = cbToggleAllExclusions.value;
        if (cbSkipLocked.enabled) {
            cbSkipLocked.value = newValue;
        }
        if (cbSkipHidden.enabled) {
            cbSkipHidden.value = newValue;
        }
        if (cbSkipSlashSlashLayers.enabled) {
            cbSkipSlashSlashLayers.value = newValue;
        }
        if (cbSkipLockedObjects.enabled) {
            cbSkipLockedObjects.value = newValue;
        }
        if (cbSkipHiddenObjects.enabled) {
            cbSkipHiddenObjects.value = newValue;
        }
        updateToggleAllExclusionsState();
    };

    cbSkipLocked.onClick = updateToggleAllExclusionsState;
    cbSkipHidden.onClick = updateToggleAllExclusionsState;
    cbSkipSlashSlashLayers.onClick = updateToggleAllExclusionsState;
    cbSkipLockedObjects.onClick = updateToggleAllExclusionsState;
    cbSkipHiddenObjects.onClick = updateToggleAllExclusionsState;
    updateToggleAllExclusionsState();

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
        skipSlashSlashLayers: cbSkipSlashSlashLayers.value,
        skipLockedObjects: cbSkipLockedObjects.value,
        skipHiddenObjects: cbSkipHiddenObjects.value,
        promoteSublayersToTopLevel: cbPromoteSublayers.value,
        guideMode: rbSeparateGuides.value ? 'separate' : (rbKeepGuidesInCurrentLayer.value ? 'keep' : 'integrate'),
        guideLayerName: normalizeLayerName(etGuideLayerName.text, '_guide'),
        mergedLayerName: etLayerName.text,
        deleteEmptyLayers: cbDeleteEmpty.value,
        reuseExistingMergedLayer: cbReuseExistingMergedLayer.value,
        layerColorValue: etLayerColor.text
    };
}

function getOrCreateMergedLayer(documentRef, mergedLayerName, reuseExistingMergedLayer) {
    var normalizedName = normalizeLayerName(mergedLayerName, '_mergedLayer');
    var mergedLayer = null;

    if (reuseExistingMergedLayer) {
        try {
            mergedLayer = documentRef.layers.getByName(normalizedName);
        } catch (e) {
            mergedLayer = null;
        }
    }

    if (mergedLayer === null) {
        mergedLayer = documentRef.layers.add();
        mergedLayer.name = reuseExistingMergedLayer ? normalizedName : generateUniqueLayerName(documentRef, normalizedName);
    }

    return mergedLayer;
}

function normalizeLayerName(layerName, fallbackName) {
    var normalized = layerName == null ? '' : String(layerName).replace(/^\s+|\s+$/g, '');
    return normalized !== '' ? normalized : fallbackName;
}

function generateUniqueLayerName(documentRef, baseName) {
    var candidate = baseName;
    var suffix = 2;

    while (layerExistsByName(documentRef, candidate)) {
        candidate = baseName + ' ' + suffix;
        suffix++;
    }

    return candidate;
}

function layerExistsByName(documentRef, layerName) {
    try {
        documentRef.layers.getByName(layerName);
        return true;
    } catch (e) {
        return false;
    }
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

    mergedLayerName = normalizeLayerName(options.mergedLayerName, '_mergedLayer');

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

    // 「現在のレイヤーに保持」時は、サブレイヤー直下のガイドを上位レイヤーへ繰り上げる。
    if (options.guideMode === 'keep') {
        hoistGuideItemsFromSublayers(documentRef, stats);
    }

    if (options.promoteSublayersToTopLevel) {
        promoteSublayersToTop(documentRef, mergedLayer, options, stats);
    }

    if (options.guideMode === 'separate') {
        separateGuidesToLayer(documentRef, mergedLayer, options, stats);
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
    if (stats.guideSeparationFailureCount > 0) {
        messages.push((lang === 'ja' ? 'ガイド分離失敗' : 'Guide separation failures') + ': ' + stats.guideSeparationFailureCount);
    }
    if (stats.parentAccessFailureCount > 0) {
        messages.push((lang === 'ja' ? '親レイヤーアクセス失敗' : 'Parent access failures') + ': ' + stats.parentAccessFailureCount);
    }
    if (messages.length > 0) {
        alert(messages.join('\n'));
    }
}

function separateGuidesToLayer(documentRef, mergedLayer, options, stats) {
    var guideLayerName = normalizeLayerName(options.guideLayerName, '_guide');
    var guideLayer = getOrCreateGuideLayer(documentRef, guideLayerName);
    applyMergedLayerSettings(guideLayer, options.layerColorValue);

    moveGuideItemsFromLayer(mergedLayer, guideLayer, stats);
}

function getOrCreateGuideLayer(documentRef, guideLayerName) {
    guideLayerName = normalizeLayerName(guideLayerName, '_guide');
    try {
        return documentRef.layers.getByName(guideLayerName);
    } catch (e) {
        var guideLayer = documentRef.layers.add();
        guideLayer.name = guideLayerName;
        return guideLayer;
    }
}

function moveGuideItemsFromLayer(sourceLayer, destinationLayer, stats) {
    var pageItems = sourceLayer.pageItems;
    for (var i = pageItems.length - 1; i >= 0; i--) {
        var item = pageItems[i];
        if (item.parent !== sourceLayer) {
            continue;
        }
        if (!isGuideItem(item)) {
            continue;
        }

        var wasLocked = false;
        try {
            wasLocked = item.locked;
            if (wasLocked) {
                item.locked = false;
            }
            item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
        } catch (e) {
            if (stats) {
                stats.guideSeparationFailureCount++;
            }
        } finally {
            try {
                if (wasLocked) {
                    item.locked = true;
                }
            } catch (restoreError) {
                if (stats) {
                    stats.itemLockRestoreFailureCount++;
                }
            }
        }
    }
}

function isGuideItem(item) {
    if (!item) {
        return false;
    }

    try {
        if (item.guides === true) {
            return true;
        }
    } catch (e) { }

    return false;
}

function hoistGuideItemsFromSublayers(documentRef, stats) {
    var topLayers = documentRef.layers;
    for (var i = 0; i < topLayers.length; i++) {
        hoistGuideItemsFromChildLayers(topLayers[i], stats);
    }
}

function hoistGuideItemsFromChildLayers(parentLayer, stats) {
    var childLayers = parentLayer.layers;
    for (var i = childLayers.length - 1; i >= 0; i--) {
        var childLayer = childLayers[i];
        hoistGuideItemsFromChildLayers(childLayer, stats);
        moveDirectGuideItemsToLayer(childLayer, parentLayer, stats);
    }
}

function moveDirectGuideItemsToLayer(sourceLayer, destinationLayer, stats) {
    var pageItems = sourceLayer.pageItems;
    for (var i = pageItems.length - 1; i >= 0; i--) {
        var item = pageItems[i];
        if (item.parent !== sourceLayer) {
            continue;
        }
        if (!isGuideItem(item)) {
            continue;
        }

        var wasLocked = false;
        try {
            wasLocked = item.locked;
            if (wasLocked) {
                item.locked = false;
            }
            item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
        } catch (e) {
            if (stats) {
                stats.guideSeparationFailureCount++;
            }
        } finally {
            try {
                if (wasLocked) {
                    item.locked = true;
                }
            } catch (restoreError) {
                if (stats) {
                    stats.itemLockRestoreFailureCount++;
                }
            }
        }
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
    if (options.skipSlashSlashLayers && isSlashSlashLayer(sourceLayer)) return false;

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
            if (options.guideMode === 'keep' && isGuideItem(item)) {
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

function isSlashSlashLayer(layerRef) {
    if (!layerRef || layerRef.name == null) {
        return false;
    }
    return String(layerRef.name).indexOf('//') === 0;
}

/* Exclude-name map for layers (case-insensitive, trimmed) */
var EXCLUDE = {
    "bg": 1
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
        if (options.skipHiddenLayers && !currentLayer.visible) {
            continue;
        }
        if (options.skipSlashSlashLayers && isSlashSlashLayer(currentLayer)) {
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
    return hasDirectPageItems(layerRef) || hasDirectGuideItems(layerRef) || layerRef.layers.length > 0;
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
// - 非表示 / ロック除外はトップレベル・サブレイヤーを問わずここで判定
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

        var skipCurrentLayerDeletion = false;

        if (options.skipLockedLayers && currentLayer.locked) {
            skipCurrentLayerDeletion = true;
        }
        if (options.skipHiddenLayers && !currentLayer.visible) {
            skipCurrentLayerDeletion = true;
        }
        if (options.skipSlashSlashLayers && isSlashSlashLayer(currentLayer)) {
            skipCurrentLayerDeletion = true;
        }

        if (currentLayer.layers.length > 0) {
            findEmptyLayers(currentLayer, resultArray, mergedLayer, options);
        }

        if (skipCurrentLayerDeletion) {
            continue;
        }

        var hasItems = hasDirectPageItems(currentLayer) || hasDirectGuideItems(currentLayer);
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

function hasDirectGuideItems(layerRef) {
    var items = layerRef.pathItems;
    for (var i = 0; i < items.length; i++) {
        if (items[i].parent === layerRef && isGuideItem(items[i])) {
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
        var parentStates = null;
        try {
            parentStates = ensureLayerParentsAccessible(layerRef, stats);
            if (!parentStates) {
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
        } finally {
            restoreLayerAccessStates(parentStates, stats);
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

    var states = [];
    for (var i = 0; i < chain.length; i++) {
        try {
            states.push({
                layer: chain[i],
                visible: chain[i].visible,
                locked: chain[i].locked
            });
            chain[i].visible = true;
            chain[i].locked = false;
        } catch (e) {
            restoreLayerAccessStates(states, stats);
            return false;
        }
    }
    return states;
}

function restoreLayerAccessStates(states, stats) {
    if (!states || states.length === 0) {
        return;
    }

    for (var i = states.length - 1; i >= 0; i--) {
        try {
            states[i].layer.visible = states[i].visible;
        } catch (e1) {
            if (stats) {
                stats.visibilityRestoreFailureCount++;
            }
        }

        try {
            states[i].layer.locked = states[i].locked;
        } catch (e2) {
            if (stats) {
                stats.layerLockRestoreFailureCount++;
            }
        }
    }
}

main();