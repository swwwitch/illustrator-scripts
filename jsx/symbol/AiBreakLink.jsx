#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択したシンボルインスタンスのリンクを解除し、通常のオブジェクトとして扱える状態にする
Illustrator 用 JSX スクリプト。

- 選択範囲内のシンボルを、グループ内も含めてまとめて処理
- static シンボル／dynamic シンボルの解除結果を整理し、扱いやすいグループまたは単体アイテムに整える
- 1アイテムだけのグループを自動的に解除するかどうかを選択可能
- 解除結果グループをネストごと完全に解除するかどうかを選択可能（`ungroupAll`）
- 解除後のアイテム名に、元のシンボル名を引き継ぐことが可能
- 名前に `symbol_unlinked_` などの接頭辞を付けるかどうかを選択可能
- 解除結果が単一テキストの場合、元のシンボル名より優先してテキスト内容を名前として使うことが可能
- 完全解除時はコンテナ名が失われるため、シンボル名系の命名オプションはダイアログ上でディム表示
- 処理後は、解除後に生成されたアイテムを選択状態として残す
- ダイアログUIとログ用メッセージは日本語・英語に対応
- エラーが起きた場合も、可能な範囲でほかのシンボルの処理を継続
- ダイアログの表示／非表示は `SHOW_OPTIONS_DIALOG` で切り替え可能

Overview

This Illustrator JSX script breaks links for selected symbol instances and turns them into regular editable objects.

- Processes selected symbols in batch, including symbols inside groups
- Organizes static and dynamic symbol break results into manageable groups or single items
- Can automatically ungroup groups that contain only one item
- Can fully ungroup the result group including nested groups (`ungroupAll`)
- Can inherit the original symbol name for the generated result
- Can optionally add a prefix such as `symbol_unlinked_` to the result name
- Can use text content as the result name when the generated result is a single text object, taking priority over the original symbol name
- Dims the symbol-name naming options in the dialog when full ungroup is on, since container names would be discarded
- Keeps the generated unlinked items selected after processing
- Supports Japanese and English dialog UI and log messages
- Continues processing other symbols as much as possible when recoverable errors occur
- `SHOW_OPTIONS_DIALOG` controls whether the options dialog is shown
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiBreakLink";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

var SHOW_OPTIONS_DIALOG = true; // 起動時にオプションダイアログを表示する / Show the options dialog on launch

var UNGROUP_SINGLE_ITEM_GROUP_DEFAULT = true; // 1アイテムだけのグループを自動的に解除する / Auto-ungroup one-item groups
var UNGROUP_ALL_DEFAULT = false; // 解除結果グループを完全に（ネストごと）解除する / Fully ungroup the result group, including nested groups
var INHERIT_SYMBOL_NAME_DEFAULT = true; // 解除後の名前に元のシンボル名を使う / Reuse the original symbol name
var USE_PREFIX_DEFAULT = true; // 解除後の名前に接頭辞を付ける / Add a prefix to the result name
var UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT = "symbol_unlinked_"; // 解除後の名前に付ける接頭辞 / Prefix for the result name
var USE_TEXT_CONTENT_AS_NAME_DEFAULT = true; // 単一テキストの場合は内容を名前にする / Use text content as name for a single text

// =========================================
// ローカライズ / Localization
// =========================================

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialog: {
        title: { ja: "シンボルのリンクを解除", en: "Break Symbol Links" }
    },
    panel: {
        options: { ja: "オプション", en: "Options" },
        resultName: { ja: "解除後の名前", en: "Result Name" }
    },
    checkbox: {
        ungroupSingleItem: { ja: "1アイテムだけのグループを解除", en: "Ungroup one-item groups" },
        ungroupAll: { ja: "グループを完全に解除", en: "Ungroup completely" },
        inheritSymbolName: { ja: "元のシンボル名を使う", en: "Use original symbol name" },
        prefix: { ja: "名前の接頭辞", en: "Name prefix" },
        useTextContentAsName: { ja: "単一テキストは内容を名前にする", en: "Use single text content as name" }
    },
    tooltip: {
        ungroupSingleItem: {
            ja: "解除後に中身が1つだけのグループができた場合、そのグループを外して単体オブジェクトにします。",
            en: "If breaking produces a group containing only one item, remove that group and keep the item standalone."
        },
        ungroupAll: {
            ja: "解除結果のグループをネストも含めて完全に解除します。グループに付けた名前は失われます。",
            en: "Fully ungroup the result, including nested groups. Any name given to a group is lost."
        },
        inheritSymbolName: {
            ja: "解除後のオブジェクト名に、元のシンボル名を引き継ぎます。",
            en: "Carry over the original symbol name to the unlinked object."
        },
        prefix: {
            ja: "シンボル名の前に付ける文字列。例: symbol_unlinked_ボタン",
            en: "Text added before the symbol name, e.g. symbol_unlinked_Button."
        },
        useTextContentAsName: {
            ja: "解除結果が1つのテキストのとき、その内容をオブジェクト名にします（シンボル名より優先）。",
            en: "When the result is a single text object, use its content as the name (takes priority over the symbol name)."
        }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" }
    },
    log: {
        moveStaticSubLayerItem: { ja: "static 解除結果のサブレイヤーアイテムを変換後グループへ移動", en: "Move static sublayer item into converted group" },
        removeConvertedSubLayer: { ja: "変換済みサブレイヤーの削除", en: "Remove converted sublayer" },
        moveSingleItemBeforeConvertedGroup: { ja: "単一アイテムを変換後グループの前へ移動", en: "Move single item before converted group" },
        removeSingleItemConvertedGroup: { ja: "単一アイテム化した変換後グループの削除", en: "Remove single-item converted group" },
        moveStaticBreakItemIntoResultGroup: { ja: "static 解除結果アイテムを結果グループへ移動", en: "Move static break item into result group" },
        selectDynamicBreakItem: { ja: "dynamic 解除結果アイテムを選択", en: "Select dynamic break item" },
        selectTargetSymbolItem: { ja: "解除対象のシンボルを選択", en: "Select target symbol item" },
        selectGeneratedResultItem: { ja: "生成された解除結果アイテムを選択", en: "Select generated break-result item" },
        ungroupAllResultGroup: { ja: "解除結果グループを完全に解除", en: "Fully ungroup the result group" },
        ungroupNestedDynamicBreakGroup: { ja: "dynamic 解除結果のネストグループを解除", en: "Ungroup nested dynamic break group" },
        moveDynamicBreakItemIntoResultGroup: { ja: "dynamic 解除結果アイテムを結果グループへ移動", en: "Move dynamic break item into result group" },
        detectBreakLinkResult: { ja: "breakLink 結果の判定", en: "Detect breakLink result" },
        mixedBreakLinkResult: {
            ja: "新規サブレイヤーと新規ページアイテムの両方が見つかりました。static として処理します。",
            en: "Both generated sublayers and page items were found. Treating the result as static."
        },
        noBreakLinkResult: {
            ja: "新規サブレイヤーも新規ページアイテムも見つかりませんでした。",
            en: "No generated sublayers or page items were found."
        },
        breakNormalizeAndNameSymbol: { ja: "シンボルのリンク解除・正規化・命名", en: "Break, normalize, and name symbol" }
    }
};

/* ドットパス（例 "panel.options"）でラベルを取得 / Get a label by dot path (e.g. "panel.options") */
function getLabel(key) {
    var parts = key.split(".");
    var entry = LABELS;
    for (var partIndex = 0; partIndex < parts.length; partIndex++) {
        if (!entry) break;
        entry = entry[parts[partIndex]];
    }
    if (entry && entry[lang]) return entry[lang];
    if (entry && entry.en) return entry.en;
    return key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLabel(key) + (lang === "ja" ? "：" : ":");
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
    var UNGROUP_ALL = UNGROUP_ALL_DEFAULT;
    var INHERIT_SYMBOL_NAME = INHERIT_SYMBOL_NAME_DEFAULT;
    var USE_PREFIX = USE_PREFIX_DEFAULT;
    var UNLINKED_SYMBOL_ITEM_NAME_PREFIX = UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT;
    var USE_TEXT_CONTENT_AS_NAME = USE_TEXT_CONTENT_AS_NAME_DEFAULT;

    /* パネルの共通レイアウトを設定 / Apply the shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    /* オプションダイアログを表示し、選択された設定を返す（キャンセル時は null）/ Show the options dialog and return the chosen settings (null on cancel) */
    function showOptionsDialog(defaultUngroupSingleItem, defaultUngroupAll, defaultInheritSymbolName, defaultUsePrefix, defaultPrefix, defaultUseTextContentAsName) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);

        var ungroupCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.ungroupSingleItem"));
        ungroupCheckbox.value = defaultUngroupSingleItem;
        ungroupCheckbox.helpTip = getLabel("tooltip.ungroupSingleItem");

        var ungroupAllCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.ungroupAll"));
        ungroupAllCheckbox.value = defaultUngroupAll;
        ungroupAllCheckbox.helpTip = getLabel("tooltip.ungroupAll");

        var namePanel = dialog.add("panel", undefined, getLabel("panel.resultName"));
        setupPanel(namePanel, 6);

        var inheritNameCheckbox = namePanel.add("checkbox", undefined, getLabel("checkbox.inheritSymbolName"));
        inheritNameCheckbox.value = defaultInheritSymbolName;
        inheritNameCheckbox.helpTip = getLabel("tooltip.inheritSymbolName");

        var prefixGroup = namePanel.add("group");
        prefixGroup.orientation = "row";
        prefixGroup.alignChildren = "center";
        var prefixCheckbox = prefixGroup.add("checkbox", undefined, labelText("checkbox.prefix"));
        prefixCheckbox.value = defaultUsePrefix;
        prefixCheckbox.helpTip = getLabel("tooltip.prefix");
        var prefixInput = prefixGroup.add("edittext", undefined, defaultPrefix);
        prefixInput.characters = 20;
        prefixInput.helpTip = getLabel("tooltip.prefix");

        var useTextNameCheckbox = namePanel.add("checkbox", undefined, getLabel("checkbox.useTextContentAsName"));
        useTextNameCheckbox.value = defaultUseTextContentAsName;
        useTextNameCheckbox.helpTip = getLabel("tooltip.useTextContentAsName");

        function syncEnabledStates() {
            var fullUngroup = ungroupAllCheckbox.value;
            /* 完全解除が ON なら 1アイテム解除は無意味なので無効化 / When full ungroup is on, the single-item option is moot, so disable it */
            ungroupCheckbox.enabled = !fullUngroup;
            /* 完全解除が ON ならコンテナに付ける名前は破棄されるため、シンボル名系の命名を無効化 / When full ungroup is on, container names are discarded, so disable the symbol-name options */
            inheritNameCheckbox.enabled = !fullUngroup;
            prefixCheckbox.enabled = !fullUngroup && inheritNameCheckbox.value;
            prefixInput.enabled = !fullUngroup && inheritNameCheckbox.value && prefixCheckbox.value;
            useTextNameCheckbox.enabled = true;
        }
        syncEnabledStates();
        inheritNameCheckbox.onClick = syncEnabledStates;
        ungroupCheckbox.onClick = syncEnabledStates;
        ungroupAllCheckbox.onClick = syncEnabledStates;
        prefixCheckbox.onClick = syncEnabledStates;

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering */
        dialog.onShow = function () {
            dialog.layout.layout(true);
            dialog.layout.resize();
        };

        if (dialog.show() !== 1) return null;

        return {
            ungroupSingleItem: ungroupCheckbox.value,
            ungroupAll: ungroupAllCheckbox.value,
            inheritSymbolName: inheritNameCheckbox.value,
            usePrefix: prefixCheckbox.value,
            unlinkedPrefix: prefixInput.text,
            useTextContentAsName: useTextNameCheckbox.value
        };
    }

    /* コンテキスト付きでエラーを $.writeln に出力 / Write an error with its context to $.writeln */
    function logScriptError(context, errorObject) {
        try {
            $.writeln("[" + SCRIPT_VERSION + "] " + context + ": " + errorObject);
        } catch (logError) {
            /* ログ出力できない環境では何もしない / Ignore when log output is unavailable */
        }
    }

    /* 以下の try* は、失敗をログして処理を続行させる安全ラッパー / The try* helpers below log failures and let processing continue */
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
            tryMoveItem(generatedSubLayer.pageItems[pageItemIndex], convertedGroup, ElementPlacement.PLACEATBEGINNING, getLabel("log.moveStaticSubLayerItem"));
        }

        tryRemoveItem(generatedSubLayer, getLabel("log.removeConvertedSubLayer"));

        if (UNGROUP_SINGLE_ITEM_GROUP && convertedGroup.pageItems.length === 1) {
            var singleItem = convertedGroup.pageItems[0];
            tryMoveItem(singleItem, convertedGroup, ElementPlacement.PLACEBEFORE, getLabel("log.moveSingleItemBeforeConvertedGroup"));
            tryRemoveItem(convertedGroup, getLabel("log.removeSingleItemConvertedGroup"));
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

    /* 設定に応じて接頭辞を付けた解除後アイテム名を組み立て / Build the result item name with the prefix when enabled */
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

    /* ExtendScript のコレクションを通常の配列へコピー / Copy an ExtendScript collection into a plain array */
    function toArray(collection) {
        var items = [];
        for (var itemIndex = 0; itemIndex < collection.length; itemIndex++) {
            items.push(collection[itemIndex]);
        }
        return items;
    }

    /* breakLink 前後の差分から新規に生成された pageItem / subLayer を抽出 / Extract newly generated pageItems or subLayers by comparing before/after references */
    function collectNewItemsByReference(currentItems, existingItems) {
        var newItems = [];

        for (var currentItemIndex = 0; currentItemIndex < currentItems.length; currentItemIndex++) {
            var currentItem = currentItems[currentItemIndex];
            var isExistingItem = false;

            for (var existingItemIndex = 0; existingItemIndex < existingItems.length; existingItemIndex++) {
                if (existingItems[existingItemIndex] === currentItem) {
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

    /* 生成物の種類から解除結果を static / dynamic / mixed / none に分類 / Classify the break result as static / dynamic / mixed / none from what was generated */
    function classifyBreakResult(generatedSubLayers, generatedPageItems) {
        var hasGeneratedSubLayers = generatedSubLayers.length > 0;
        var hasGeneratedPageItems = generatedPageItems.length > 0;

        if (hasGeneratedSubLayers && !hasGeneratedPageItems) return "static";
        if (!hasGeneratedSubLayers && hasGeneratedPageItems) return "dynamic";
        if (hasGeneratedSubLayers && hasGeneratedPageItems) return "mixed";
        return "none";
    }

    /* シンボルインスタンスのリンクを解除 / Break the link of the symbol instance */
    function breakSymbolLink(symbolItem) {
        symbolItem.breakLink();
    }

    /* 分類結果に応じて static / dynamic の整理関数へ振り分け / Route to the static or dynamic organizer based on the classification */
    function normalizeBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers, generatedPageItems) {
        var breakResultType = classifyBreakResult(generatedSubLayers, generatedPageItems);

        if (breakResultType === "static") {
            return organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers);
        }

        if (breakResultType === "dynamic") {
            return organizeDynamicBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak);
        }

        if (breakResultType === "mixed") {
            logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.mixedBreakLinkResult"));
            return organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, generatedSubLayers);
        }

        logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.noBreakLinkResult"));
        return null;
    }

    /* 解除結果アイテムに名前を付与（単一テキスト内容 → 元シンボル名の順で優先）/ Name the result item, preferring single-text content over the original symbol name */
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

    /* 選択範囲全体から SymbolItem をまとめて収集 / Collect all SymbolItems across the whole selection */
    function collectSymbolItemsFromSelection(selectedItems) {
        var symbolItems = [];

        for (var selectionIndex = 0; selectionIndex < selectedItems.length; selectionIndex++) {
            collectSymbolItemsRecursively(selectedItems[selectionIndex], symbolItems);
        }

        return symbolItems;
    }

    /* static 解除結果（サブレイヤー生成型）をグループ／単体に整理 / Organize a static break result (sublayer-based) into a group or single item */
    function organizeStaticBreakResult(targetLayer, existingLayerItemsBeforeBreak, newlyCreatedSubLayers) {
        if (!targetLayer) return null;

        for (var subLayerIndex = 0; subLayerIndex < newlyCreatedSubLayers.length; subLayerIndex++) {
            convertGeneratedSubLayerToGroup(targetLayer, newlyCreatedSubLayers[subLayerIndex]);
        }

        var newlyCreatedItems = collectNewItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak);

        if (newlyCreatedItems.length > 1) {
            var resultGroup = targetLayer.groupItems.add();
            for (var newItemIndex = 0; newItemIndex < newlyCreatedItems.length; newItemIndex++) {
                tryMoveItem(newlyCreatedItems[newItemIndex], resultGroup, ElementPlacement.PLACEATEND, getLabel("log.moveStaticBreakItemIntoResultGroup"));
            }
            return resultGroup;
        }

        if (newlyCreatedItems.length === 1) {
            return newlyCreatedItems[0];
        }

        return null;
    }

    /* dynamic 解除結果（ページアイテム生成型）を ungroup で平坦化しグループ／単体に整理 / Organize a dynamic break result (pageItem-based) by ungrouping, into a group or single item */
    function organizeDynamicBreakResult(documentObject, targetLayer, existingLayerItemsBeforeBreak) {
        if (!targetLayer) return null;

        /* breakLink 直後の selection に依存せず、Layer 配下の差分で新規アイテムを特定 / Detect generated items by Layer reference differences instead of relying on the selection after breakLink */
        var newlyCreatedItems = collectNewItemsByReference(targetLayer.pageItems, existingLayerItemsBeforeBreak);
        if (newlyCreatedItems.length === 0) return null;

        documentObject.selection = null;
        for (var newItemIndex = 0; newItemIndex < newlyCreatedItems.length; newItemIndex++) {
            trySetItemSelected(newlyCreatedItems[newItemIndex], true, getLabel("log.selectDynamicBreakItem"));
        }

        /* ネストが深い場合に備え、ungroup の反復回数に上限を設ける / Cap the ungroup iterations to guard against deeply nested groups */
        var ungroupIterationGuard = 0;
        while (selectionContainsGroupItem(documentObject.selection) && ungroupIterationGuard < 50) {
            if (!tryExecuteMenuCommand("ungroup", getLabel("log.ungroupNestedDynamicBreakGroup"))) break;
            ungroupIterationGuard++;
        }

        if (documentObject.selection.length > 1) {
            var dynamicTargetLayer = getContainingLayer(documentObject.selection[0]) || targetLayer;
            if (!dynamicTargetLayer) return null;

            /* selection は移動中に変化するため、事前に配列へ退避 / Snapshot the selection into an array since it changes during moves */
            var selectedItems = toArray(documentObject.selection);

            var resultGroup = dynamicTargetLayer.groupItems.add();
            for (var selectedItemIndex = 0; selectedItemIndex < selectedItems.length; selectedItemIndex++) {
                tryMoveItem(selectedItems[selectedItemIndex], resultGroup, ElementPlacement.PLACEATEND, getLabel("log.moveDynamicBreakItemIntoResultGroup"));
            }

            return resultGroup;
        }

        if (documentObject.selection.length === 1) {
            return documentObject.selection[0];
        }

        return null;
    }

    /* breakLink 前の Layer 直下の pageItems / subLayers を控えておく / Snapshot the Layer's direct pageItems and subLayers before breakLink */
    function snapshotLayerContents(targetLayer) {
        return {
            pageItems: targetLayer ? toArray(targetLayer.pageItems) : [],
            subLayers: targetLayer ? toArray(targetLayer.layers) : []
        };
    }

    /* 解除結果グループをネストごと完全に解除し、展開後アイテムの配列を返す / Fully ungroup the result group and return the resulting loose items */
    function ungroupResultGroupCompletely(documentObject, resultItem) {
        if (!resultItem) return [];
        if (resultItem.typename !== "GroupItem") return [resultItem];

        documentObject.selection = null;
        if (!trySetItemSelected(resultItem, true, getLabel("log.ungroupAllResultGroup"))) return [resultItem];

        /* ungroupAll はネストごと一括で解除する / ungroupAll dissolves the group and all nested groups at once */
        tryExecuteMenuCommand("ungroupAll", getLabel("log.ungroupAllResultGroup"));

        /* 解除直後の選択が展開後アイテム / The selection right after ungroupAll holds the loose items */
        return toArray(documentObject.selection);
    }

    /* SymbolItem ごとの処理フローをまとめ、選択に残す結果アイテムの配列を返す / Coordinate the per-symbol flow and return the result items to keep selected */
    function breakNormalizeAndNameSymbol(symbolItem, documentObject) {
        if (symbolItem.locked || symbolItem.hidden) return [];
        /* 外側で初期選択を解除済みの前提で、この SymbolItem だけを処理対象として選択 / Select only this SymbolItem, assuming the outer flow already cleared the initial selection */
        if (!trySetItemSelected(symbolItem, true, getLabel("log.selectTargetSymbolItem"))) return [];

        var symbolName = symbolItem.symbol.name;
        var targetLayer = getContainingLayer(symbolItem);
        var beforeBreak = snapshotLayerContents(targetLayer);

        breakSymbolLink(symbolItem);

        var newlyCreatedSubLayers = targetLayer ? collectNewItemsByReference(targetLayer.layers, beforeBreak.subLayers) : [];
        var newlyCreatedLayerItems = targetLayer ? collectNewItemsByReference(targetLayer.pageItems, beforeBreak.pageItems) : [];

        var finalOutputItem = normalizeBreakResult(
            documentObject,
            targetLayer,
            beforeBreak.pageItems,
            newlyCreatedSubLayers,
            newlyCreatedLayerItems
        );

        applyNaming(finalOutputItem, symbolName);

        /* 完全解除が ON なら、命名後に結果グループをネストごと解除して展開後アイテムを返す / When full ungroup is on, dissolve the result group after naming and return the loose items */
        if (UNGROUP_ALL) {
            return ungroupResultGroupCompletely(documentObject, finalOutputItem);
        }

        return finalOutputItem ? [finalOutputItem] : [];
    }

    var symbolItems = collectSymbolItemsFromSelection(activeDocument.selection);

    if (symbolItems.length === 0) {
        return;
    }

    if (SHOW_OPTIONS_DIALOG) {
        var dialogResult = showOptionsDialog(
            UNGROUP_SINGLE_ITEM_GROUP_DEFAULT,
            UNGROUP_ALL_DEFAULT,
            INHERIT_SYMBOL_NAME_DEFAULT,
            USE_PREFIX_DEFAULT,
            UNLINKED_SYMBOL_ITEM_NAME_PREFIX_DEFAULT,
            USE_TEXT_CONTENT_AS_NAME_DEFAULT
        );
        if (!dialogResult) {
            return;
        }
        UNGROUP_SINGLE_ITEM_GROUP = dialogResult.ungroupSingleItem;
        UNGROUP_ALL = dialogResult.ungroupAll;
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
    /* 初期選択はここで一度だけ解除し、生成物を配列に集めて末尾でまとめて選択する / Clear the initial selection once here, collect generated items, and select them all at the end */
    activeDocument.selection = null;

    var generatedResultItems = [];
    for (var symbolIndex = 0; symbolIndex < symbolItems.length; symbolIndex++) {
        try {
            var producedItems = breakNormalizeAndNameSymbol(symbolItems[symbolIndex], activeDocument);
            for (var producedIndex = 0; producedIndex < producedItems.length; producedIndex++) {
                generatedResultItems.push(producedItems[producedIndex]);
            }
        } catch (errorObject) {
            logScriptError(getLabel("log.breakNormalizeAndNameSymbol"), errorObject);
        }
    }

    /* 全シンボルの生成物を最終的に選択状態へ（static / dynamic / 完全解除で共通）/ Select every symbol's generated items at the end (uniform for static / dynamic / full-ungroup) */
    activeDocument.selection = null;
    for (var resultIndex = 0; resultIndex < generatedResultItems.length; resultIndex++) {
        trySetItemSelected(generatedResultItems[resultIndex], true, getLabel("log.selectGeneratedResultItem"));
    }

})();