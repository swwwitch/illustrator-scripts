#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    スクリプト名：SuperLayerManage.jsx

    概要: 
    Illustrator ドキュメント内のオブジェクトを指定レイヤーへ一括移動します。
    対象は「選択オブジェクト」「すべてのオブジェクト」「すべてのテキストオブジェクト」から選べます。
    オプションで空レイヤーの削除が可能です（"bg" と "//" で始まるレイヤーは削除対象外）。

    English Description:
    Move objects in an Illustrator document to a specified layer in bulk.
    You can choose between "Selected Objects", "All Objects", or "All Text Objects".
    Optionally delete empty layers (layers named "bg" or starting with "//" are excluded).

    更新履歴：
    - v1.0.0（2025-07-03） : 初版リリース
*/

// Unlock, unhide, and move item
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
    for (var i = 0; i < layers.length; i++) layerNames[i] = layers[i].name;

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

        // Move selected objects
        if (radioSelected.value) {
            var sel = doc.selection;
            if (!sel || sel.length === 0) {
                alert(LABELS.noSelection[lang]);
                return;
            }
            for (var i = 0; i < sel.length; i++) {
                prepareAndMoveItem(sel[i], targetLayer);
            }

        // Move all text objects
        } else if (radioAllText.value) {
            var allTextItems = doc.textFrames;
            for (var j = 0; j < allTextItems.length; j++) {
                prepareAndMoveItem(allTextItems[j], targetLayer);
            }

        // Move all objects
        } else {
            app.executeMenuCommand('unlockAll');
            app.executeMenuCommand('showAll');

            // Collect all items recursively from all layers
            var allItems = [];
            for (var l = 0; l < doc.layers.length; l++) {
                collectAllItemsRecursive(doc.layers[l], allItems);
            }

            for (var k = 0; k < allItems.length; k++) {
                prepareAndMoveItem(allItems[k], targetLayer);
            }
        }

        // Optionally delete empty layers except "bg" and layers starting with "//"
        if (deleteEmptyLayersCheckbox.value) {
            for (var l = layers.length - 1; l >= 0; l--) {
                var lyr = layers[l];
                var name = lyr.name;
                if (lyr.pageItems.length === 0 && name !== "bg" && name.indexOf("//") !== 0) {
                    lyr.remove();
                }
            }
        }

        dlg.close();
    };

    closeBtn.onClick = function() {
        dlg.close();
    };

    dlg.show();
}

main();
// Recursively collect all items (including inside groups) into arr
function collectAllItemsRecursive(container, arr) {
    for (var i = 0; i < container.pageItems.length; i++) {
        var item = container.pageItems[i];
        arr.push(item);
        if (item.typename === "GroupItem") {
            collectAllItemsRecursive(item, arr);
        }
    }
}