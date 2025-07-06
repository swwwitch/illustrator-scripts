#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 


/*
### スクリプト名：

SortItemsByPosition.jsx

### 概要

- ページアイテムをX座標またはY座標で並べ替え、重ね順を更新します。
- 視覚的な整理やレイアウト確認を効率化できます。

### 主な機能

- X/Y座標による昇順ソート
- 反転ボタンで順序の切り替え
- 選択オブジェクトまたはアートボード内オブジェクト限定選択
- レイヤーを移動し、空レイヤーを削除するオプション

### 処理の流れ

1. 対象オブジェクトを収集
2. 並べ替え基準を選択
3. プレビューを確認
4. OKで確定

### 更新履歴

- v1.0.0 (20250706) : 初期バージョン

---

### Script Name:

SortItemsByPosition.jsx

### Overview

- Sort page items by X or Y coordinate and update stacking order.
- Helps efficiently organize and visually check layouts.

### Main Features

- Ascending sort by X or Y coordinate
- Reverse button to toggle order
- Limit to selection or current artboard objects
- Option to move objects to top layer and delete empty layers

### Workflow

1. Collect target objects
2. Select sort criteria
3. Preview order
4. Confirm with OK

### Changelog

- v1.0.0 (20250706) : Initial version
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "重ね順の変更 v1.0", en: "Reorder Objects v1.0" },
    sortPanel:   { ja: "ソート基準", en: "Sort Criteria" },
    xLeft:       { ja: "X座標（左右）", en: "X (Horizontal)" },
    yTop:        { ja: "Y座標（上下）", en: "Y (Vertical)" },
    targetPanel: { ja: "対象", en: "Target" },
    selection:   { ja: "選択オブジェクト", en: "Selection" },
    artboard:    { ja: "現在のアートボードに限定", en: "Limit to Current Artboard" },
    moveLayer:   { ja: "レイヤーを移動", en: "Move to Top Layer" },
    ok:          { ja: "OK", en: "OK" },
    reverse:     { ja: "反転", en: "Reverse" },
    cancel:      { ja: "キャンセル", en: "Cancel" }
};

function main() {
    var doc = app.activeDocument;

    /*
      アートボード内判定
      item: PageItem, abBounds: [left, top, right, bottom]
    */
    function isInsideArtboard(item, abBounds) {
        var pos = item.position;
        var left = pos[0];
        var top = pos[1];
        return (left >= abBounds[0] && left <= abBounds[2] && top <= abBounds[1] && top >= abBounds[3]);
    }

    /*
      対象オブジェクト収集
      useSelection: 選択のみ, useArtboard: アートボード内のみ
    */
    function collectItems(doc, useSelection, useArtboard) {
        var result = [];
        if (useSelection && doc.selection.length > 0) {
            for (var i = 0; i < doc.selection.length; i++) {
                result.push(doc.selection[i]);
            }
            return result;
        }

        if (useArtboard) {
            var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            var abBounds = artboard.artboardRect;
            for (var i = 0; i < doc.pageItems.length; i++) {
                var item = doc.pageItems[i];
                if (isInsideArtboard(item, abBounds)) {
                    result.push(item);
                }
            }
        } else {
            for (var i = 0; i < doc.pageItems.length; i++) {
                result.push(doc.pageItems[i]);
            }
        }
        return result;
    }

    /* 配列シャッフル */
    function shuffleArray(array) {
        var m = array.length, t, i;
        while (m) {
            i = Math.floor(Math.random() * m--);
            t = array[m];
            array[m] = array[i];
            array[i] = t;
        }
    }

    /*
      並べ替えを適用
      mode: "yAsc"=Y座標昇順(下が背面), "yDesc"=Y座標降順(上を背面にする), "xAsc"=X座標昇順(左が背面)
    */
    function applySorting(items, mode) {
        if (mode === "yAsc") {
            // Y座標（下から）：下が背面になるよう昇順
            items.sort(function(a, b) { return a.position[1] - b.position[1]; });
        } else if (mode === "yDesc") {
            // Y座標（上を背面にするための降順）
            items.sort(function(a, b) { return b.position[1] - a.position[1]; });
        } else if (mode === "random") {
            shuffleArray(items);
        } else {
            // X座標（左から昇順：左が背面）
            items.sort(function(a, b) { return a.position[0] - b.position[0]; });
        }
    }

    /* 重ね順を更新 */
    function updateZOrder(items) {
        for (var i = 0; i < items.length; i++) {
            items[i].zOrder(ZOrderMethod.BRINGTOFRONT);
        }
        app.redraw();
    }

    /* プレビュー反映 */
    function applyPreview(items, mode) {
        applySorting(items, mode);
        updateZOrder(items);
    }

    var hasSelection = (doc.selection.length > 0);

    /*
      選択オブジェクトの配置方向によって自動的にXまたはYモードを設定
    */
    var autoMode = "x";
    if (hasSelection && doc.selection.length > 1) {
        var minX = doc.selection[0].geometricBounds[0];
        var maxX = doc.selection[0].geometricBounds[2];
        var minY = doc.selection[0].geometricBounds[1];
        var maxY = doc.selection[0].geometricBounds[3];
        for (var i = 1; i < doc.selection.length; i++) {
            var b = doc.selection[i].geometricBounds;
            if (b[0] < minX) minX = b[0];
            if (b[2] > maxX) maxX = b[2];
            if (b[1] > minY) minY = b[1];
            if (b[3] < maxY) maxY = b[3];
        }
        var width = maxX - minX;
        var height = minY - maxY;
        if (height >= width) {
            autoMode = "y";
        }
    }

    /* itemsの初期取得（ダイアログ構築より前に移動） */
    var items = collectItems(doc, hasSelection, !hasSelection);

    /* ダイアログ構築開始 */
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    /* メイングループ（2カラムレイアウト） */
    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = "top";

    /* 左カラム */
    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = "left";

    /* ソートパネル */
    var radioPanel = leftGroup.add("panel", undefined, LABELS.sortPanel[lang]);
    radioPanel.orientation = "column";
    radioPanel.alignChildren = "left";
    radioPanel.margins = [15, 20, 15, 10];

    var sortRadioXLeft = radioPanel.add("radiobutton", undefined, LABELS.xLeft[lang]);
    var sortRadioYTop = radioPanel.add("radiobutton", undefined, LABELS.yTop[lang]);
    if (autoMode === "y") {
        sortRadioYTop.value = true;
    } else {
        sortRadioXLeft.value = true;
    }

    /* 対象パネル */
    var targetPanel = leftGroup.add("panel", undefined, LABELS.targetPanel[lang]);
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 10];
    var selectionRadio = targetPanel.add("radiobutton", undefined, LABELS.selection[lang]);
    var artboardRadio = targetPanel.add("radiobutton", undefined, LABELS.artboard[lang]);
    selectionRadio.value = hasSelection;
    artboardRadio.value = !hasSelection;

    /* 左カラム内の2つのパネルの幅を統一 */
    var uniformPanelWidth = 200;
    radioPanel.preferredSize.width = uniformPanelWidth;
    targetPanel.preferredSize.width = uniformPanelWidth;

    /* レイヤー移動チェックボックス用グループ */
    var moveLayerGroup = leftGroup.add("group");
    moveLayerGroup.orientation = "column";
    moveLayerGroup.alignChildren = "center";
    moveLayerGroup.preferredSize.width = uniformPanelWidth;
    var moveTopLayerCheckbox = moveLayerGroup.add("checkbox", undefined, LABELS.moveLayer[lang]);
    moveTopLayerCheckbox.value = false;

    // 選択オブジェクトが単一レイヤーの場合、レイヤー移動チェックボックスを無効化
    var uniqueLayers = {};
    for (var i = 0; i < items.length; i++) {
        uniqueLayers[items[i].layer.name] = true;
    }
    var layerKeys = [];
    for (var key in uniqueLayers) {
        layerKeys.push(key);
    }
    if (layerKeys.length <= 1) {
        moveTopLayerCheckbox.enabled = false;
    }

    /* 右カラム（ボタングループ） */
    var rightGroup = mainGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = "right";
    var buttonWidth = 100;
    var okBtn = rightGroup.add("button", undefined, LABELS.ok[lang], { name: "OK" });
    okBtn.preferredSize.width = buttonWidth;

    var reverseBtn = rightGroup.add("button", undefined, LABELS.reverse[lang]);
    reverseBtn.preferredSize.width = buttonWidth;

    var spacer = rightGroup.add("statictext", undefined, "");
    spacer.preferredSize.height = 100;

    var cancelBtn = rightGroup.add("button", undefined, LABELS.cancel[lang]);
    cancelBtn.preferredSize.width = buttonWidth;


    /* 反転ボタン有効/無効の更新（常に有効） */
    function updateReverseButtonState() {
        reverseBtn.enabled = true;
    }


    /* プレビュー共通イベント処理 */
    function previewWithMode(mode) {
        applyPreview(items, mode);
        updateReverseButtonState();
    }

    /* ソート基準変更時のプレビュー更新 */
    sortRadioXLeft.onClick = function() {
        previewWithMode("xAsc");
    };
    sortRadioYTop.onClick = function() {
        previewWithMode("yDesc");
    };

    /* 現在のソートモード取得 */
    function getCurrentMode() {
        return sortRadioYTop.value ? "yDesc" : "xAsc";
    }

    /* 対象切り替え時のプレビュー更新 */
    selectionRadio.onClick = function() {
        items = collectItems(doc, selectionRadio.value === true, artboardRadio.value === true);
        previewWithMode(getCurrentMode());
    };
    artboardRadio.onClick = selectionRadio.onClick;

    /* 反転ボタン処理 */
    reverseBtn.onClick = function() {
        items.reverse();
        updateZOrder(items);
    };

    updateReverseButtonState();

    if (dialog.show() == 1) {
        items = collectItems(doc, selectionRadio.value === true, artboardRadio.value === true);
        applySorting(items, getCurrentMode());
        updateZOrder(items);
        /* レイヤーを移動チェック時の処理 */
        if (moveTopLayerCheckbox.value) {
            var topLayer = doc.layers[0];
            // 移動前に逆順にして順序を保持
            items.reverse();
            for (var i = 0; i < items.length; i++) {
                items[i].move(topLayer, ElementPlacement.PLACEATEND);
            }
            // 移動後に重ね順を更新して見た目を保持
            updateZOrder(items);
            // 元レイヤーが空なら削除
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                if (layer.pageItems.length === 0) {
                    layer.remove();
                }
            }
        }
    }
    /* ダイアログ後に items をクリア */
    items = [];
}

main();