#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択したシンボルインスタンスのリンクを解除し、通常のオブジェクトとして扱える状態にする
Illustrator 用 JSX スクリプト。

- 選択範囲内のシンボルを、グループ内も含めてまとめて処理
- static シンボル／dynamic シンボルの解除結果を整理し、扱いやすいグループまたは単体アイテムに整える
- 解除後のアイテム名に、元のシンボル名を引き継ぐことが可能
- 名前に `symbol_unlinked_` などの接頭辞を付けるかどうかを選択可能
- 解除結果が単一テキストの場合、元のシンボル名より優先してテキスト内容を名前として使うことが可能
- 1アイテムだけのグループを自動的に解除するかどうかを選択可能
- 処理後は、解除後に生成されたアイテムを選択状態として残す
- ダイアログUIとログ用メッセージは日本語・英語に対応
- エラーが起きた場合も、可能な範囲でほかのシンボルの処理を継続
- ダイアログの表示／非表示は `SHOW_OPTIONS_DIALOG` で切り替え可能

Overview

This Illustrator JSX script breaks links for selected symbol instances and turns them into regular editable objects.

- Processes selected symbols in batch, including symbols inside groups
- Organizes static and dynamic symbol break results into manageable groups or single items
- Can inherit the original symbol name for the generated result
- Can optionally add a prefix such as `symbol_unlinked_` to the result name
- Can use text content as the result name when the generated result is a single text object, taking priority over the original symbol name
- Can automatically ungroup groups that contain only one item
- Keeps the generated unlinked items selected after processing
- Supports Japanese and English dialog UI and log messages
- Continues processing other symbols as much as possible when recoverable errors occur
- `SHOW_OPTIONS_DIALOG` controls whether the options dialog is shown
*/

/* =========================================
   設定 / Settings
   ========================================= */

var SHOW_OPTIONS_DIALOG = true; // 起動時にオプションダイアログを表示する

var UNGROUP_SINGLE_ITEM_GROUP_DEFAULT = true; // 1アイテムだけのグループを自動的に解除する
var INHERIT_SYMBOL_NAME_DEFAULT = true; // 解除後の名前に元のシンボル名を使う
var USE_PREFIX_DEFAULT = true; // 解除後の名前に接頭辞を付ける
var UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT = "symbol_unlinked_"; // 解除後の名前に付ける接頭辞
var USE_TEXT_CONTENT_AS_NAME_DEFAULT = true; // 単一テキストの場合は内容を名前にする

/* =========================================
   バージョンとローカライズ / Version and localization
   ========================================= */

var SCRIPT_VERSION = "v1.1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    title: {
        ja: "シンボルのリンクを解除",
        en: "Break Symbol Links"
    },
    optionsPanelTitle: {
        ja: "オプション",
        en: "Options"
    },
    ungroupSingleItem: {
        ja: "1アイテムだけのグループを解除",
        en: "Ungroup one-item groups"
    },
    namePanelTitle: {
        ja: "解除後の名前",
        en: "Result Name"
    },
    inheritSymbolName: {
        ja: "元のシンボル名を使う",
        en: "Use original symbol name"
    },
    prefix: {
        ja: "名前の接頭辞",
        en: "Name prefix"
    },
    useTextContentAsName: {
        ja: "単一テキストは内容を名前にする",
        en: "Use single text content as name"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    logMoveStaticSubLayerItem: {
        ja: "static 解除結果のサブレイヤーアイテムを変換後グループへ移動",
        en: "Move static sublayer item into converted group"
    },
    logRemoveConvertedSubLayer: {
        ja: "変換済みサブレイヤーの削除",
        en: "Remove converted sublayer"
    },
    logMoveSingleItemBeforeConvertedGroup: {
        ja: "単一アイテムを変換後グループの前へ移動",
        en: "Move single item before converted group"
    },
    logRemoveSingleItemConvertedGroup: {
        ja: "単一アイテム化した変換後グループの削除",
        en: "Remove single-item converted group"
    },
    logMoveStaticBreakItemIntoResultGroup: {
        ja: "static 解除結果アイテムを結果グループへ移動",
        en: "Move static break item into result group"
    },
    logSelectDynamicBreakItem: {
        ja: "dynamic 解除結果アイテムを選択",
        en: "Select dynamic break item"
    },
    logSelectTargetSymbolItem: {
        ja: "解除対象のシンボルを選択",
        en: "Select target symbol item"
    },
    logUngroupNestedDynamicBreakGroup: {
        ja: "dynamic 解除結果のネストグループを解除",
        en: "Ungroup nested dynamic break group"
    },
    logMoveDynamicBreakItemIntoResultGroup: {
        ja: "dynamic 解除結果アイテムを結果グループへ移動",
        en: "Move dynamic break item into result group"
    },
    logDetectBreakLinkResult: {
        ja: "breakLink 結果の判定",
        en: "Detect breakLink result"
    },
    logMixedBreakLinkResult: {
        ja: "新規サブレイヤーと新規ページアイテムの両方が見つかりました。static として処理します。",
        en: "Both generated sublayers and page items were found. Treating the result as static."
    },
    logNoBreakLinkResult: {
        ja: "新規サブレイヤーも新規ページアイテムも見つかりませんでした。",
        en: "No generated sublayers or page items were found."
    },
    logBreakNormalizeAndNameSymbol: {
        ja: "シンボルのリンク解除・正規化・命名",
        en: "Break, normalize, and name symbol"
    }
};

function L(key) {
    if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    return key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}


(function () {

    if (app.documents.length === 0) {
        return;
    }

    var activeDocument = app.activeDocument;

    if (activeDocument.selection.length === 0) {
        return;
    }

    var PANEL_MARGINS = [15, 20, 15, 10];

    var UNGROUP_SINGLE_ITEM_GROUP = UNGROUP_SINGLE_ITEM_GROUP_DEFAULT;
    var INHERIT_SYMBOL_NAME = INHERIT_SYMBOL_NAME_DEFAULT;
    var USE_PREFIX = USE_PREFIX_DEFAULT;
    var UNLINKED_SYMBOL_ITEM_NAME_PREFIX = UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT;
    var USE_TEXT_CONTENT_AS_NAME = USE_TEXT_CONTENT_AS_NAME_DEFAULT;

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    function showOptionsDialog(defaultUngroupSingleItem, defaultInheritSymbolName, defaultUsePrefix, defaultPrefix, defaultUseTextContentAsName) {
        var dialog = new Window("dialog", L("title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        var optionsPanel = dialog.add("panel", undefined, L("optionsPanelTitle"));
        setupPanel(optionsPanel, 6);

        var ungroupCheckbox = optionsPanel.add("checkbox", undefined, L("ungroupSingleItem"));
        ungroupCheckbox.value = defaultUngroupSingleItem;

        var namePanel = dialog.add("panel", undefined, L("namePanelTitle"));
        setupPanel(namePanel, 6);

        var inheritNameCheckbox = namePanel.add("checkbox", undefined, L("inheritSymbolName"));
        inheritNameCheckbox.value = defaultInheritSymbolName;

        var prefixGroup = namePanel.add("group");
        prefixGroup.orientation = "row";
        prefixGroup.alignChildren = "center";
        var prefixCheckbox = prefixGroup.add("checkbox", undefined, labelText("prefix"));
        prefixCheckbox.value = defaultUsePrefix;
        var prefixInput = prefixGroup.add("edittext", undefined, defaultPrefix);
        prefixInput.characters = 20;

        var useTextNameCheckbox = namePanel.add("checkbox", undefined, L("useTextContentAsName"));
        useTextNameCheckbox.value = defaultUseTextContentAsName;

        function syncEnabledStates() {
            prefixCheckbox.enabled = inheritNameCheckbox.value;
            prefixInput.enabled = inheritNameCheckbox.value && prefixCheckbox.value;
            useTextNameCheckbox.enabled = true;
        }
        syncEnabledStates();
        inheritNameCheckbox.onClick = syncEnabledStates;
        ungroupCheckbox.onClick = syncEnabledStates;
        prefixCheckbox.onClick = syncEnabledStates;

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, L("ok"), { name: "ok" });

        if (dialog.show() !== 1) return null;

        return {
            ungroupSingleItem: ungroupCheckbox.value,
            inheritSymbolName: inheritNameCheckbox.value,
            usePrefix: prefixCheckbox.value,
            unlinkedPrefix: prefixInput.text,
            useTextContentAsName: useTextNameCheckbox.value
        };
    }

    function logScriptError(context, errorObject) {
        try {
            $.writeln("[" + SCRIPT_VERSION + "] " + context + ": " + errorObject);
        } catch (logError) {
            /* ログ出力できない環境では何もしない / Ignore when log output is unavailable */
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

    /* 指定サブレイヤーを親 Layer 直下のグループに変換 / Convert the specified sublayer into a group directly under the parent Layer */
    /* 単一アイテムのグループは設定に応じて解除 / Ungroup single-item groups depending on the option */
    function convertGeneratedSubLayerToGroup(parentLayer, generatedSubLayer) {
        if (!parentLayer || !generatedSubLayer) return null;

        flattenSubLayersToGroups(generatedSubLayer);

        var convertedGroup = parentLayer.groupItems.add();
        for (var pageItemIndex = generatedSubLayer.pageItems.length - 1; pageItemIndex >= 0; pageItemIndex--) {
            tryMoveItem(generatedSubLayer.pageItems[pageItemIndex], convertedGroup, ElementPlacement.PLACEATBEGINNING, L("logMoveStaticSubLayerItem"));
        }

        tryRemoveItem(generatedSubLayer, L("logRemoveConvertedSubLayer"));

        if (UNGROUP_SINGLE_ITEM_GROUP && convertedGroup.pageItems.length === 1) {
            var singleItem = convertedGroup.pageItems[0];
            tryMoveItem(singleItem, convertedGroup, ElementPlacement.PLACEBEFORE, L("logMoveSingleItemBeforeConvertedGroup"));
            tryRemoveItem(convertedGroup, L("logRemoveSingleItemConvertedGroup"));
            return singleItem;
        }

        return convertedGroup;
    }

    /* 指定 Layer 配下のサブレイヤーを再帰的にグループ化してフラット化 / Flatten sublayers under the specified Layer by converting them into groups recursively */
    function flattenSubLayersToGroups(parentLayer) {
        for (var subLayerIndex = parentLayer.layers.length - 1; subLayerIndex >= 0; subLayerIndex--) {
            convertGeneratedSubLayerToGroup(parentLayer, parentLayer.layers[subLayerIndex]);
        }
    }

    /* 選択範囲内に GroupItem が含まれるか / Check whether the selection contains any GroupItem */
    function selectionContainsGroupItem(selectedItems) {
        for (var selectionIndex = 0; selectionIndex < selectedItems.length; selectionIndex++) {
            if (selectedItems[selectionIndex].typename === "GroupItem") return true;
        }
        return false;
    }

    /* 指定アイテムを含む Layer を返す / Return the Layer that contains the specified item */
    function getContainingLayer(pageItem) {
        var currentParent = pageItem.parent;
        while (currentParent && currentParent.typename !== "Layer") {
            currentParent = currentParent.parent;
        }
        return currentParent;
    }

    function getUnlinkedSymbolItemName(symbolName) {
        return (USE_PREFIX ? UNLINKED_SYMBOL_ITEM_NAME_PREFIX : "") + symbolName;
    }

    /* 単一 TextFrame の文字列を取得 / Get the text content when the result is a single TextFrame */
    /* 改行は空白に置換し、対象外の場合は null を返す / Replace line breaks with spaces and return null when not applicable */
    function getSingleTextFrameContent(item) {
        if (!item) return null;

        var textFrame = null;
        if (item.typename === "TextFrame") {
            textFrame = item;
        } else if (item.typename === "GroupItem" && item.pageItems.length === 1 && item.pageItems[0].typename === "TextFrame") {
            textFrame = item.pageItems[0];
        }

        if (!textFrame) return null;

        var contents = textFrame.contents;
        if (contents === null || typeof contents === "undefined") return null;

        contents = contents.replace(/[\r\n]+/g, " ");
        if (contents.length === 0) return null;

        return contents;
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

    function classifyBreakResult(generatedSubLayers, generatedPageItems) {
        var hasGeneratedSubLayers = generatedSubLayers.length > 0;
        var hasGeneratedPageItems = generatedPageItems.length > 0;

        if (hasGeneratedSubLayers && !hasGeneratedPageItems) return "static";
        if (!hasGeneratedSubLayers && hasGeneratedPageItems) return "dynamic";
        if (hasGeneratedSubLayers && hasGeneratedPageItems) return "mixed";
        return "none";
    }

    function breakSymbolLink(symbolItem) {
        symbolItem.breakLink();
    }

    function normalizeBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers, generatedPageItems) {
        var breakResultType = classifyBreakResult(generatedSubLayers, generatedPageItems);

        if (breakResultType === "static") {
            return organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers);
        }

        if (breakResultType === "dynamic") {
            return organizeDynamicBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak);
        }

        if (breakResultType === "mixed") {
            logScriptError(L("logDetectBreakLinkResult"), L("logMixedBreakLinkResult"));
            return organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers);
        }

        logScriptError(L("logDetectBreakLinkResult"), L("logNoBreakLinkResult"));
        return null;
    }

    function applyNaming(finalOutputItem, symbolName) {
        if (!finalOutputItem) return;

        if (USE_TEXT_CONTENT_AS_NAME) {
            var textContent = getSingleTextFrameContent(finalOutputItem);
            if (textContent !== null) {
                finalOutputItem.name = textContent;
                return;
            }
        }

        if (INHERIT_SYMBOL_NAME) {
            finalOutputItem.name = getUnlinkedSymbolItemName(symbolName);
        }
    }

    /* 選択範囲から SymbolItem を収集 / Collect SymbolItems from the selection */
    function collectSymbolItemsRecursively(pageItem, symbolItems) {
        if (!pageItem) return;

        if (pageItem.typename === "SymbolItem") {
            symbolItems.push(pageItem);
            return;
        }

        /* グループ内も探す / Search inside groups as well */
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
                tryMoveItem(newlyCreatedItems[newItemIndex], resultGroup, ElementPlacement.PLACEATEND, L("logMoveStaticBreakItemIntoResultGroup"));
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

        /* breakLink 直後の selection に依存せず、Layer 配下の差分で新規アイテムを特定 / Detect generated items by Layer reference differences instead of relying on the selection after breakLink */
        var newlyCreatedItems = collectGeneratedPageItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak);
        if (newlyCreatedItems.length === 0) return null;

        documentObject.selection = null;
        for (var newItemIndex = 0; newItemIndex < newlyCreatedItems.length; newItemIndex++) {
            trySetItemSelected(newlyCreatedItems[newItemIndex], true, L("logSelectDynamicBreakItem"));
        }

        var safety = 0;
        while (selectionContainsGroupItem(documentObject.selection) && safety < 50) {
            if (!tryExecuteMenuCommand("ungroup", L("logUngroupNestedDynamicBreakGroup"))) break;
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
                tryMoveItem(selectedItems[selectedItemIndex], resultGroup, ElementPlacement.PLACEATEND, L("logMoveDynamicBreakItemIntoResultGroup"));
            }

            return resultGroup;
        }

        if (documentObject.selection.length === 1) {
            return documentObject.selection[0];
        }

        return null;
    }

    /* SymbolItem ごとの処理フローをまとめる / Coordinate the per-symbol flow: break link, classify, normalize, and name */
    function breakNormalizeAndNameSymbol(symbolItem, documentObject) {
        if (symbolItem.locked || symbolItem.hidden) return;
        /* 外側で初期選択を解除済みの前提で、この SymbolItem だけを処理対象として選択 / Select only this SymbolItem, assuming the outer flow already cleared the initial selection */
        if (!trySetItemSelected(symbolItem, true, L("logSelectTargetSymbolItem"))) return;

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
        breakSymbolLink(symbolItem);

        var newlyCreatedSubLayers = targetLayer ? collectGeneratedSubLayersByReference(targetLayer.layers, existingSubLayersBeforeBreak) : [];
        var newlyCreatedLayerItems = targetLayer ? collectGeneratedPageItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak) : [];

        var finalOutputItem = normalizeBreakResult(
            documentObject,
            targetLayer,
            existingLayerItemsBeforeBreak,
            newlyCreatedSubLayers,
            newlyCreatedLayerItems
        );

        applyNaming(finalOutputItem, symbolName);
    }

    var symbolItems = collectSymbolItemsFromSelection(activeDocument.selection);

    if (symbolItems.length === 0) {
        return;
    }

    if (SHOW_OPTIONS_DIALOG) {
        var dialogResult = showOptionsDialog(
            UNGROUP_SINGLE_ITEM_GROUP_DEFAULT,
            INHERIT_SYMBOL_NAME_DEFAULT,
            USE_PREFIX_DEFAULT,
            UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT,
            USE_TEXT_CONTENT_AS_NAME_DEFAULT
        );
        if (!dialogResult) {
            return;
        }
        UNGROUP_SINGLE_ITEM_GROUP = dialogResult.ungroupSingleItem;
        INHERIT_SYMBOL_NAME = dialogResult.inheritSymbolName;
        USE_PREFIX = dialogResult.usePrefix;
        UNLINKED_SYMBOL_ITEM_NAME_PREFIX = dialogResult.unlinkedPrefix;
        USE_TEXT_CONTENT_AS_NAME = dialogResult.useTextContentAsName;
    } else {
        /* ダイアログ非表示時は単一テキスト名を優先 / When dialog is hidden, prioritize text content as name */
        USE_TEXT_CONTENT_AS_NAME = true;
    }

    /* 処理後の選択方針 / Post-processing selection policy */
    /* 元の選択は復元せず、解除後に生成されたアイテムを選択状態として残す / Do not restore the original selection; keep the generated unlinked items selected */
    /* 初期選択はここで一度だけ解除し、dynamic 整理時のみ生成アイテムを選び直す / Clear the initial selection only once here; reselect generated items only while organizing dynamic results */
    activeDocument.selection = null;

    for (var symbolIndex = 0; symbolIndex < symbolItems.length; symbolIndex++) {
        try {
            breakNormalizeAndNameSymbol(symbolItems[symbolIndex], activeDocument);
        } catch (errorObject) {
            logScriptError(L("logBreakNormalizeAndNameSymbol"), errorObject);
        }
    }

})();