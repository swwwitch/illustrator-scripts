#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択したシンボルインスタンスのリンクを解除し、解除後に生成されたアイテムを整理・リネームする
Illustrator 用 JSX スクリプト。

- 選択範囲内の SymbolItem をグループ内も含めて再帰的に収集
- breakLink() 前後の参照差分で static 系／dynamic 系の解除結果を判定
- static 系の解除結果では、新規生成サブレイヤーだけをグループ化して整理
- dynamic 系の解除結果では、新規生成アイテムを選択してグループ解除し、必要に応じて再グループ化
- 解除後のグループまたは単体アイテムに `symbol_unlinked_` 接頭辞付きの名前を設定
- 処理後は元の選択を復元せず、解除後に生成されたアイテムを選択状態として残す
- 失敗しやすい move/remove/selection/menu command はログ化し、処理単位で失敗を分離
- グループ内のアイテムが 1 つだけの場合にグループ解除するかどうかを `UNGROUP_SINGLE_ITEM_GROUP` で切り替え可能

This Illustrator JSX script breaks links for selected symbol instances, organizes the generated result,
and renames the resulting group or single item using the `symbol_unlinked_<symbol name>` pattern.
It detects static and dynamic breakLink results by comparing object references before and after breakLink(),
keeps only the newly generated unlinked items selected after processing, and logs recoverable failures.
*/

var SCRIPT_VERSION = "v1.0.0";

(function () {

    if (app.documents.length === 0) {
        return;
    }

    var activeDocument = app.activeDocument;

    if (activeDocument.selection.length === 0) {
        return;
    }

    var UNGROUP_SINGLE_ITEM_GROUP = true;
    var UNLINKED_SYMBOL_ITEM_NAME_PREFIX = "symbol_unlinked_";

    function logScriptError(context, errorObject) {
        try {
            $.writeln("[" + SCRIPT_VERSION + "] " + context + ": " + errorObject);
        } catch (logError) {
            // $.writeln が使えない環境では何もしない / Ignore when $.writeln is unavailable
        }
    }

    function tryRemoveItem(item, context) {
        try {
            item.remove();
            return true;
        } catch (errorObject) {
            logScriptError(context, errorObject);
            return false;
        }
    }

    function tryMoveItem(item, destination, placement, context) {
        try {
            item.move(destination, placement);
            return true;
        } catch (errorObject) {
            logScriptError(context, errorObject);
            return false;
        }
    }

    function trySetItemSelected(item, selectedValue, context) {
        try {
            item.selected = selectedValue;
            return true;
        } catch (errorObject) {
            logScriptError(context, errorObject);
            return false;
        }
    }

    function tryExecuteMenuCommand(commandName, context) {
        try {
            app.executeMenuCommand(commandName);
            return true;
        } catch (errorObject) {
            logScriptError(context, errorObject);
            return false;
        }
    }

    // 指定サブレイヤーを親 Layer 直下のグループに変換する
    // 変換後のグループの子が 1 つだけの場合、設定に応じて解除して中身を親 Layer 直下に出す
    function convertGeneratedSubLayerToGroup(parentLayer, generatedSubLayer) {
        if (!parentLayer || !generatedSubLayer) return null;

        convertNestedSubLayersToGroups(generatedSubLayer);

        var convertedGroup = parentLayer.groupItems.add();
        for (var pageItemIndex = generatedSubLayer.pageItems.length - 1; pageItemIndex >= 0; pageItemIndex--) {
            tryMoveItem(generatedSubLayer.pageItems[pageItemIndex], convertedGroup, ElementPlacement.PLACEATBEGINNING, "move static sublayer item into converted group");
        }

        tryRemoveItem(generatedSubLayer, "remove converted sublayer");

        if (UNGROUP_SINGLE_ITEM_GROUP && convertedGroup.pageItems.length === 1) {
            var singleItem = convertedGroup.pageItems[0];
            tryMoveItem(singleItem, convertedGroup, ElementPlacement.PLACEBEFORE, "move single item before converted group");
            tryRemoveItem(convertedGroup, "remove single item converted group");
            return singleItem;
        }

        return convertedGroup;
    }

    // 指定 Layer 配下のサブレイヤーをそれぞれ Layer 直下のグループに変換する（再帰）
    function convertNestedSubLayersToGroups(parentLayer) {
        for (var subLayerIndex = parentLayer.layers.length - 1; subLayerIndex >= 0; subLayerIndex--) {
            convertGeneratedSubLayerToGroup(parentLayer, parentLayer.layers[subLayerIndex]);
        }
    }

    // 選択範囲内に GroupItem が含まれるか
    function selectionContainsGroupItem(selectedItems) {
        for (var selectionIndex = 0; selectionIndex < selectedItems.length; selectionIndex++) {
            if (selectedItems[selectionIndex].typename === "GroupItem") return true;
        }
        return false;
    }

    // 指定アイテムの祖先 Layer を返す
    function getContainingLayer(pageItem) {
        var currentParent = pageItem.parent;
        while (currentParent && currentParent.typename !== "Layer") {
            currentParent = currentParent.parent;
        }
        return currentParent;
    }

    function getUnlinkedSymbolItemName(symbolName) {
        return UNLINKED_SYMBOL_ITEM_NAME_PREFIX + symbolName;
    }

    function collectGeneratedPageItemsByReference(currentPageItems, existingPageItems) {
        var newItems = [];

        for (var currentItemIndex = 0; currentItemIndex < currentPageItems.length; currentItemIndex++) {
            var currentItem = currentPageItems[currentItemIndex];
            var isExistingItem = false;

            for (var existingItemIndex = 0; existingItemIndex < existingPageItems.length; existingItemIndex++) {
                if (existingPageItems[existingItemIndex] === currentItem) {
                    isExistingItem = true;
                    break;
                }
            }

            if (!isExistingItem) {
                newItems.push(currentItem);
            }
        }

        return newItems;
    }

    function collectGeneratedSubLayersByReference(currentSubLayers, existingSubLayers) {
        var newLayers = [];

        for (var currentLayerIndex = 0; currentLayerIndex < currentSubLayers.length; currentLayerIndex++) {
            var currentLayer = currentSubLayers[currentLayerIndex];
            var isExistingLayer = false;

            for (var existingLayerIndex = 0; existingLayerIndex < existingSubLayers.length; existingLayerIndex++) {
                if (existingSubLayers[existingLayerIndex] === currentLayer) {
                    isExistingLayer = true;
                    break;
                }
            }

            if (!isExistingLayer) {
                newLayers.push(currentLayer);
            }
        }

        return newLayers;
    }

    // 選択範囲から SymbolItem を収集
    function collectSymbolItemsRecursively(pageItem, symbolItems) {
        if (!pageItem) return;

        if (pageItem.typename === "SymbolItem") {
            symbolItems.push(pageItem);
            return;
        }

        // グループ内も探す
        if (pageItem.typename === "GroupItem") {
            for (var pageItemIndex = 0; pageItemIndex < pageItem.pageItems.length; pageItemIndex++) {
                collectSymbolItemsRecursively(pageItem.pageItems[pageItemIndex], symbolItems);
            }
        }
    }

    function collectSymbolItemsFromSelection(selectedItems) {
        var symbolItems = [];

        for (var selectionIndex = 0; selectionIndex < selectedItems.length; selectionIndex++) {
            collectSymbolItemsRecursively(selectedItems[selectionIndex], symbolItems);
        }

        return symbolItems;
    }

    function organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, newlyCreatedSubLayers) {
        if (!targetLayer) return null;

        for (var subLayerIndex = 0; subLayerIndex < newlyCreatedSubLayers.length; subLayerIndex++) {
            convertGeneratedSubLayerToGroup(targetLayer, newlyCreatedSubLayers[subLayerIndex]);
        }

        var newlyCreatedItems = collectGeneratedPageItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak);

        if (newlyCreatedItems.length > 1) {
            var resultGroup = targetLayer.groupItems.add();
            for (var newItemIndex = 0; newItemIndex < newlyCreatedItems.length; newItemIndex++) {
                tryMoveItem(newlyCreatedItems[newItemIndex], resultGroup, ElementPlacement.PLACEATEND, "move static break item into result group");
            }
            return resultGroup;
        }

        if (newlyCreatedItems.length === 1) {
            return newlyCreatedItems[0];
        }

        return null;
    }

    function organizeDynamicBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak) {
        if (!targetLayer) return null;

        // breakLink 直後は selection が当てにならないため、Layer 配下の差分で新規アイテムを特定する
        var newlyCreatedItems = collectGeneratedPageItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak);
        if (newlyCreatedItems.length === 0) return null;

        documentObject.selection = null;
        for (var newItemIndex = 0; newItemIndex < newlyCreatedItems.length; newItemIndex++) {
            trySetItemSelected(newlyCreatedItems[newItemIndex], true, "select dynamic break item");
        }

        tryExecuteMenuCommand("ungroupAll", "ungroup all dynamic break items");

        var safety = 0;
        while (selectionContainsGroupItem(documentObject.selection) && safety < 50) {
            if (!tryExecuteMenuCommand("ungroup", "ungroup nested dynamic break group")) break;
            safety++;
        }

        if (documentObject.selection.length > 1) {
            var dynamicTargetLayer = getContainingLayer(documentObject.selection[0]) || targetLayer;
            if (!dynamicTargetLayer) return null;

            var selectedItems = [];
            for (var selectionIndex = 0; selectionIndex < documentObject.selection.length; selectionIndex++) {
                selectedItems.push(documentObject.selection[selectionIndex]);
            }

            var resultGroup = dynamicTargetLayer.groupItems.add();
            for (var selectedItemIndex = 0; selectedItemIndex < selectedItems.length; selectedItemIndex++) {
                tryMoveItem(selectedItems[selectedItemIndex], resultGroup, ElementPlacement.PLACEATEND, "move dynamic break item into result group");
            }

            return resultGroup;
        }

        if (documentObject.selection.length === 1) {
            return documentObject.selection[0];
        }

        return null;
    }

    function breakSymbolLinkAndNameResult(symbolItem, documentObject) {
        if (symbolItem.locked || symbolItem.hidden) return;

        documentObject.selection = null;
        symbolItem.selected = true;

        var symbolName = symbolItem.symbol.name;
        var targetLayer = getContainingLayer(symbolItem);
        var existingLayerItemsBeforeBreak = [];
        var existingSubLayersBeforeBreak = [];
        if (targetLayer) {
            for (var preItemIndex = 0; preItemIndex < targetLayer.pageItems.length; preItemIndex++) {
                existingLayerItemsBeforeBreak.push(targetLayer.pageItems[preItemIndex]);
            }
            for (var preLayerIndex = 0; preLayerIndex < targetLayer.layers.length; preLayerIndex++) {
                existingSubLayersBeforeBreak.push(targetLayer.layers[preLayerIndex]);
            }
        }

        symbolItem.breakLink();

        var newlyCreatedSubLayers = targetLayer ? collectGeneratedSubLayersByReference(targetLayer.layers, existingSubLayersBeforeBreak) : [];
        var newlyCreatedLayerItems = targetLayer ? collectGeneratedPageItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak) : [];
        var isStaticBreakResult = newlyCreatedSubLayers.length > 0;
        var isDynamicBreakResult = newlyCreatedLayerItems.length > 0;

        var finalOutputItem = null;

        if (isStaticBreakResult) {
            finalOutputItem = organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, newlyCreatedSubLayers);
        } else if (isDynamicBreakResult) {
            finalOutputItem = organizeDynamicBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak);
        } else {
            logScriptError("detect breakLink result", "No generated sublayers or page items were found.");
        }

        if (finalOutputItem) {
            finalOutputItem.name = getUnlinkedSymbolItemName(symbolName);
        }
    }

    var symbolItems = collectSymbolItemsFromSelection(activeDocument.selection);

    if (symbolItems.length === 0) {
        return;
    }

    // 処理後の選択方針 / Post-processing selection policy
    // 元の選択は復元せず、解除後に生成されたアイテムを選択状態として残す
    // Do not restore the original selection; keep the generated unlinked items selected.
    // 解除対象を個別処理するため、ここで元の選択を解除する
    activeDocument.selection = null;

    for (var symbolIndex = 0; symbolIndex < symbolItems.length; symbolIndex++) {
        try {
            breakSymbolLinkAndNameResult(symbolItems[symbolIndex], activeDocument);
        } catch (errorObject) {
            logScriptError("process symbol break", errorObject);
        }
    }

})();