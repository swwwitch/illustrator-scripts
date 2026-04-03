#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddTrimMark.jsx

### 概要

- 選択したオブジェクト、または現在のアートボードを対象にトリムマークを作成するIllustrator用スクリプトです。
- 実行時にダイアログを表示し、対象を「選択したオブジェクト」または「現在のアートボード」から選べます。
- トリムマークは専用の「トンボ」レイヤーに配置され、必要に応じてガイドも自動生成されます。

### 主な機能

- ダイアログで処理対象を選択（選択オブジェクト / 現在のアートボード）
- 選択オブジェクトを基にトリムマークを作成
- 現在のアートボード全体を基にトリムマークを作成
- トリムマークを「トンボ」レイヤーに移動し、自動ロック
- 元オブジェクトのガイド化

### 処理の流れ

1. ダイアログで「選択したオブジェクト」または「現在のアートボード」を選択
2. 対象オブジェクトを複製し、塗り・線をなしに設定
3. トリムマーク作成メニューを実行後、複製オブジェクトを削除
4. トリムマークを「トンボ」レイヤーに移動してロック
5. 選択オブジェクトを対象にした場合のみ、元オブジェクトを複製してガイド化

### 更新履歴

- v1.0 (20250205) : 初期バージョン
- v1.1 (20250603) : コメント整理と処理安定化
- v1.1.5 (20260401) : 作業用レイヤー削除前にアクティブレイヤーを戻すよう修正し、空の作業用レイヤーを正常に削除

---

### Script Name:

AddTrimMark.jsx

### Overview

- An Illustrator script that creates trim marks for either selected objects or the current artboard.
- When launched, it shows a dialog that lets you choose the target: "Selected Objects" or "Current Artboard".
- Trim marks are placed on a dedicated "Trim" layer, and guides are created when needed.

### Main Features

- Choose the processing target in a dialog (Selected Objects / Current Artboard)
- Create trim marks based on selected objects
- Create trim marks based on the current artboard
- Move trim marks to a dedicated "Trim" layer and lock it automatically
- Convert the original object to a guide when using selected objects

### Process Flow

1. Choose either "Selected Objects" or "Current Artboard" in the dialog
2. Duplicate the target object and remove fill and stroke
3. Execute the trim mark command, then remove the duplicated object
4. Move the trim marks to the "Trim" layer and lock it
5. Only when selected objects are used, duplicate the original object and convert it to a guide

### Update History

- v1.0 (20250205): Initial version
- v1.1 (20250603): Refined comments and stabilized process
- v1.1.5 (20260401): Restored the active layer before deleting the working layer so empty working layers are removed correctly
*/

function showTargetDialog(hasSelection) {
    var dlg = new Window('dialog', 'トリムマークを追加');
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';
    dlg.margins = [15, 20, 15, 15];

    var pnl = dlg.add('panel', undefined, '対象');
    pnl.orientation = 'column';
    pnl.alignChildren = 'left';
    pnl.margins = [15, 20, 15, 10];

    var rbSelection = pnl.add('radiobutton', undefined, '選択したオブジェクト');
    var rbArtboard = pnl.add('radiobutton', undefined, '現在のアートボード');

    if (hasSelection) {
        rbSelection.value = true;
    } else {
        rbSelection.enabled = false;
        rbArtboard.value = true;
    }

    var btns = dlg.add('group');
    btns.alignment = 'right';
    btns.add('button', undefined, 'キャンセル', { name: 'cancel' });
    btns.add('button', undefined, 'OK', { name: 'ok' });

    if (dlg.show() !== 1) {
        return null;
    }

    return rbSelection.value ? 'selection' : 'artboard';
}

function getOrCreateTrimLayer(doc) {
    var trimLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === 'トンボ') {
            trimLayer = doc.layers[i];
            break;
        }
    }
    if (!trimLayer) {
        trimLayer = doc.layers.add();
        trimLayer.name = 'トンボ';
    }
    trimLayer.locked = false;
    trimLayer.visible = true;
    return trimLayer;
}

function getOrCreateWorkingLayer(doc) {
    var layerName = '__AddTrimMark_Work__';
    var workLayer = null;
    var i = 0;

    for (i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            workLayer = doc.layers[i];
            break;
        }
    }
    if (!workLayer) {
        workLayer = doc.layers.add();
        workLayer.name = layerName;
    }

    workLayer.locked = false;
    workLayer.visible = true;
    doc.activeLayer = workLayer;
    return workLayer;
}

function createInvisibleArtboardRect(doc, artboardIndex, workLayer) {
    var rect = doc.artboards[artboardIndex].artboardRect;
    var item = workLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
    item.filled = false;
    item.stroked = false;
    return item;
}

function main() {
    if (app.documents.length === 0) {
        alert('ドキュメントを開いてください。');
        return;
    }

    var doc = app.activeDocument;
    var hasSelection = doc.selection.length > 0;
    var targetMode = showTargetDialog(hasSelection);
    var trimLayer = null;
    var j = 0;
    var workLayer = null;
    var originalActiveLayer = doc.activeLayer;
    var originalSelection = [];

    if (!targetMode) {
        return;
    }

    for (j = 0; j < doc.selection.length; j++) {
        originalSelection.push(doc.selection[j]);
    }

    function moveSelectionToTrimLayer() {
        trimLayer.locked = false;
        trimLayer.visible = true;
        for (j = doc.selection.length - 1; j >= 0; j--) {
            doc.selection[j].move(trimLayer, ElementPlacement.PLACEATBEGINNING);
        }
        trimLayer.locked = true;
    }

    try {
        trimLayer = getOrCreateTrimLayer(doc);
        workLayer = getOrCreateWorkingLayer(doc);

        if (targetMode === 'selection') {
            var targetObj = originalSelection[0];
            var duplicatedObj = targetObj.duplicate();
            duplicatedObj.filled = false;
            duplicatedObj.stroked = false;

            doc.activeLayer = workLayer;
            doc.selection = [duplicatedObj];
            app.executeMenuCommand('TrimMark v25');
            duplicatedObj.remove();

            moveSelectionToTrimLayer();

            doc.selection = [targetObj];
            var guideObj = targetObj.duplicate();
            guideObj.guides = true;
        } else {
            var currentIndex = doc.artboards.getActiveArtboardIndex();
            var currentRectItem = createInvisibleArtboardRect(doc, currentIndex, workLayer);

            doc.activeLayer = workLayer;
            doc.selection = [currentRectItem];
            app.executeMenuCommand('TrimMark v25');
            currentRectItem.remove();

            moveSelectionToTrimLayer();
            doc.selection = null;
        }
    } finally {
        try {
            doc.activeLayer = originalActiveLayer;
        } catch (e) {
        }

        try {
            doc.selection = null;
            for (j = 0; j < originalSelection.length; j++) {
                try {
                    originalSelection[j].selected = true;
                } catch (e) {
                }
            }
        } catch (e) {
        }

        if (workLayer) {
            try {
                if (workLayer.pageItems.length === 0) {
                    workLayer.remove();
                } else {
                    workLayer.locked = false;
                    workLayer.visible = false;
                }
            } catch (e) {
            }
        }
    }
}

main();