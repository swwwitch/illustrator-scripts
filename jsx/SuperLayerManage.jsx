#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    スクリプト名：SuperLayerManage.jsx

    概要:
    Illustrator ドキュメント内のオブジェクトを指定レイヤーに一括移動するスクリプトです。
    「選択オブジェクト」「すべてのオブジェクト」「すべてのテキストオブジェクト」から移動対象を選べます。
    オプションで空レイヤーを削除できます（"bg" および "//" で始まるレイヤーは削除対象外）。
    移動後、対象レイヤーのカラーを RGB(79, 128, 255) に設定します。

    English Description:
    Move objects in an Illustrator document to a specified layer in bulk.
    You can choose "Selected Objects", "All Objects", or "All Text Objects".
    Optionally delete empty layers (except those named "bg" or starting with "//").
    After moving, the layer color is set to RGB(79, 128, 255).

    更新履歴：
    - v1.0.0（2025-07-03）: 初版リリース
    - v1.0.1（2025-07-03） : レイヤーカラーを変更
*/

// Unlock and unhide all objects in the document
function unlockAndShowAll() {
    app.executeMenuCommand('unlockAll');
    app.executeMenuCommand('showAll');
}

// Unlock, unhide, and move a single item to the target layer
function prepareAndMoveItem(item, targetLayer) {
    try {
        item.locked = false;
        item.hidden = false;
        item.move(targetLayer, ElementPlacement.PLACEATEND);
    } catch (e) {}
}

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var layers = doc.layers;
    var layerNames = [];
    for (var i = 0; i < layers.length; i++) {
        layerNames[i] = layers[i].name;
    }

    var LABELS = {
        dialogTitle: { ja: "レイヤー管理ツール", en: "Layer Management Tool" },
        panelTitle: { ja: "対象オブジェクト", en: "Target Objects" },
        selectedObj: { ja: "選択中", en: "Selected" },
        allObj: { ja: "すべて", en: "All Objects" },
        allText: { ja: "テキストのみ", en: "All Text" },
        deleteEmpty: { ja: "空レイヤーを削除", en: "Delete Empty Layers" },
        layerList: { ja: "移動先のレイヤー", en: "Target Layer" },
        move: { ja: "移動", en: "Move" },
        close: { ja: "閉じる", en: "Close" },
        noLayerSelected: { ja: "移動先のレイヤーを選択してください。", en: "Please select a target layer." },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." }
    };

    var lang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "row";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 20;

    var leftGroup = dlg.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["left", "top"];

    var objPanel = leftGroup.add("panel", undefined, LABELS.panelTitle[lang]);
    objPanel.orientation = "column";
    objPanel.alignChildren = ["left", "top"];
    objPanel.margins = [15, 20, 15, 10];

    var radioSelected = objPanel.add("radiobutton", undefined, LABELS.selectedObj[lang]);
    var radioAll = objPanel.add("radiobutton", undefined, LABELS.allObj[lang]);
    var radioAllText = objPanel.add("radiobutton", undefined, LABELS.allText[lang]);
    radioSelected.value = true;

    var deleteGroup = leftGroup.add("group");
    deleteGroup.orientation = "row";
    deleteGroup.alignChildren = ["center", "center"];
    deleteGroup.margins = [5, 5, 0, 0];

    var deleteEmptyLayersCheckbox = deleteGroup.add("checkbox", undefined, LABELS.deleteEmpty[lang]);
    deleteEmptyLayersCheckbox.value = true;

    var rightGroup = dlg.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = ["fill", "top"];

    var radioLayerGroup = rightGroup.add("panel", undefined, LABELS.layerList[lang]);
    radioLayerGroup.orientation = "column";
    radioLayerGroup.alignChildren = ["left", "top"];
    radioLayerGroup.margins = [15, 20, 15, 10];

    var radioButtons = [];
    for (var i = 0; i < layerNames.length; i++) {
        radioButtons[i] = radioLayerGroup.add("radiobutton", undefined, layerNames[i]);
    }
    if (radioButtons.length > 0) {
        radioButtons[0].value = true;
    }

    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignChildren = ["fill", "top"];

    var moveBtn = buttonGroup.add("button", undefined, LABELS.move[lang], {name: "ok"});
    var closeBtn = buttonGroup.add("button", undefined, LABELS.close[lang]);

    moveBtn.onClick = function() {
        var targetLayerName = null;
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                targetLayerName = radioButtons[i].text;
                break;
            }
        }
        if (!targetLayerName) {
            alert(LABELS.noLayerSelected[lang]);
            return;
        }
        var targetLayer = doc.layers.getByName(targetLayerName);

        var itemsToMove = [];

        if (radioSelected.value) {
            var sel = doc.selection;
            if (!sel || sel.length === 0) {
                alert(LABELS.noSelection[lang]);
                return;
            }
            itemsToMove = sel;
        } else if (radioAllText.value) {
            itemsToMove = doc.textFrames;
        } else {
            unlockAndShowAll();
            // Collect all pageItems recursively from all layers
            for (var l = 0; l < doc.layers.length; l++) {
                collectAllItemsRecursive(doc.layers[l], itemsToMove);
            }
        }

        for (var k = 0; k < itemsToMove.length; k++) {
            prepareAndMoveItem(itemsToMove[k], targetLayer);
        }

        // Delete empty layers except those named "bg" or starting with "//"
        if (deleteEmptyLayersCheckbox.value) {
            for (var l = layers.length - 1; l >= 0; l--) {
                var lyr = layers[l];
                var name = lyr.name;
                if (lyr.pageItems.length === 0 && name !== "bg" && name.indexOf("//") !== 0) {
                    lyr.remove();
                }
            }
        }

        // Set the target layer color to RGB(79, 128, 255)
        changeSelectedLayerColorToRGB(targetLayer, 79, 128, 255);

        dlg.close();
    };

    closeBtn.onClick = function() {
        dlg.close();
    };

    dlg.show();
}

// Recursively collect all pageItems (including inside groups) into arr
function collectAllItemsRecursive(container, arr) {
    for (var i = 0; i < container.pageItems.length; i++) {
        var item = container.pageItems[i];
        arr.push(item);
        if (item.typename === "GroupItem") {
            collectAllItemsRecursive(item, arr);
        }
    }
}

// Change the color of one or more layers to the specified RGB color
function changeSelectedLayerColorToRGB(targetLayers, r, g, b) {
    if (!targetLayers) {
        alert("レイヤーが指定されていません。");
        return;
    }
    if (targetLayers.typename === "Layer") {
        targetLayers = [targetLayers];
    }
    var newColor = new RGBColor();
    newColor.red = r;
    newColor.green = g;
    newColor.blue = b;
    for (var i = 0; i < targetLayers.length; i++) {
        var layer = targetLayers[i];
        if (layer.visible && !layer.locked) {
            layer.color = newColor;
        }
    }
}

main();