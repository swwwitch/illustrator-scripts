// Illustratorスクリプトとして実行（必須）
#target illustrator
// JSX外部スクリプト警告を非表示
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    スクリプト名：SuperLayerManage.jsx

    概要:
    Illustratorドキュメント内のオブジェクトを指定レイヤーへ一括移動するツール。
    モード: 「選択オブジェクト」「すべてのオブジェクト」「すべてのテキストオブジェクト」
    オプション: 空レイヤー削除（"bg" および "//" で始まるレイヤーは除外）
    移動先レイヤーのカラーを RGB(79, 128, 255) に変更

    更新履歴：
    - v1.0.0（2025-07-03）: 初版リリース
    - v1.0.1（2025-07-03）: レイヤーカラー変更機能追加
    - v1.0.2（2025-07-03）: 自動選択判定、空レイヤー削除ロジック改善
*/

// すべてのロックと非表示を解除
function unlockAndShowAll() {
    app.executeMenuCommand('unlockAll');
    app.executeMenuCommand('showAll');
}

// 個別アイテムのロック解除・表示・移動
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

    // UIラベル定義（未使用項目は削除済み）
    var LABELS = {
        dialogTitle: { ja: "レイヤー管理ツール", en: "Layer Management Tool" },
        panelTitle: { ja: "対象オブジェクト", en: "Target Objects" },
        selectedObj: { ja: "選択オブジェクト", en: "Selected Objects" },
        allObj: { ja: "すべてのオブジェクト", en: "All Objects" },
        allText: { ja: "テキストオブジェクト", en: "Text Only" },
        deleteEmpty: { ja: "空レイヤーを削除", en: "Delete Empty Layers" },
        layerList: { ja: "移動先レイヤー", en: "Target Layer" },
        move: { ja: "移動", en: "Move" },
        close: { ja: "閉じる", en: "Close" },
        noLayerSelected: { ja: "移動先レイヤーを選択してください。", en: "Please select a target layer." },
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
    var sel = doc.selection;
    var hasSelection = sel && sel.length > 0;

    // デフォルト選択：オブジェクト選択有無で自動判定
    radioSelected.value = hasSelection;
    radioAll.value = !hasSelection;

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
    var firstAvailableIndex = -1;
    for (var i = 0; i < layerNames.length; i++) {
        radioButtons[i] = radioLayerGroup.add("radiobutton", undefined, layerNames[i]);
        if (layers[i].locked) {
            radioButtons[i].enabled = false;
        } else if (firstAvailableIndex === -1) {
            firstAvailableIndex = i;
        }
    }
    if (firstAvailableIndex !== -1) {
        radioButtons[firstAvailableIndex].value = true;
    }

    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignChildren = ["fill", "top"];

    var moveBtn = buttonGroup.add("button", undefined, LABELS.move[lang], {name: "ok"});
    var closeBtn = buttonGroup.add("button", undefined, LABELS.close[lang]);

    moveBtn.onClick = function() {
        // 移動先レイヤー取得
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

        // 移動対象アイテム収集
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
            for (var l = 0; l < doc.layers.length; l++) {
                collectAllItemsFromLayerRecursive(doc.layers[l], itemsToMove);
            }
        }

        // アイテムを移動
        for (var k = 0; k < itemsToMove.length; k++) {
            prepareAndMoveItem(itemsToMove[k], targetLayer);
        }

        // 空レイヤー削除（bg, // 除外）
        if (deleteEmptyLayersCheckbox.value) {
            for (var l = layers.length - 1; l >= 0; l--) {
                deleteEmptyLayersRecursive(layers[l]);
            }
        }

        // レイヤーカラー変更
        changeSelectedLayerColorToRGB(targetLayer, 79, 128, 255);
        dlg.close();
    };

    closeBtn.onClick = function() {
        dlg.close();
    };

    dlg.show();
}

// コンテナ（グループなど）から全pageItemsを再帰的に収集
function collectAllItemsRecursive(container, arr) {
    for (var i = 0; i < container.pageItems.length; i++) {
        var item = container.pageItems[i];
        arr.push(item);
        if (item.typename === "GroupItem") {
            collectAllItemsRecursive(item, arr);
        }
    }
}

// レイヤー・サブレイヤーから全pageItemsを収集
function collectAllItemsFromLayerRecursive(layer, arr) {
    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        arr.push(item);
        if (item.typename === "GroupItem") {
            collectAllItemsRecursive(item, arr);
        }
    }
    for (var j = 0; j < layer.layers.length; j++) {
        collectAllItemsFromLayerRecursive(layer.layers[j], arr);
    }
}

// 空レイヤーを再帰的に削除（bg, // で始まるレイヤーは除外）
function deleteEmptyLayersRecursive(layer) {
    for (var i = layer.layers.length - 1; i >= 0; i--) {
        deleteEmptyLayersRecursive(layer.layers[i]);
    }
    if (layer.pageItems.length === 0 && layer.layers.length === 0 && layer.name !== "bg" && layer.name.indexOf("//") !== 0) {
        layer.remove();
    }
}

// レイヤーカラーをRGBで変更
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