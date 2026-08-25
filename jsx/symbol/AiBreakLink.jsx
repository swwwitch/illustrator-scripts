#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したシンボルインスタンスのリンクを解除し、通常のオブジェクトとして扱える状態にします。

詳細は README を参照してください。

### Overview

Breaks the links of the selected symbol instances and turns them into regular editable objects.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiBreakLink";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiBreakLink.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiBreakLink.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf729c53f4300"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 起動時にオプションダイアログを表示する / Show the options dialog on launch */
var SHOW_OPTIONS_DIALOG = true;

/* 1アイテムだけのグループを自動的に解除する / Auto-ungroup one-item groups */
var UNGROUP_SINGLE_ITEM_GROUP_DEFAULT = true;

/* 解除結果グループを完全に（ネストごと）解除する / Fully ungroup the result group, including nested groups */
var UNGROUP_ALL_DEFAULT = false;

/* 解除後の名前に元のシンボル名を使う / Reuse the original symbol name */
var INHERIT_SYMBOL_NAME_DEFAULT = true;

/* 解除後の名前に接頭辞を付ける / Add a prefix to the result name */
var USE_PREFIX_DEFAULT = true;

/* 解除後の名前に付ける接頭辞 / Prefix for the result name */
var UNLINKED_ITEM_NAME_PREFIX_DEFAULT = "symbol_unlinked_";

/* 単一テキストの場合は内容を名前にする / Use text content as name for a single text */
var USE_TEXT_CONTENT_AS_NAME_DEFAULT = true;

/* ネストしたグループを解除する際の反復回数の上限 / Cap on ungroup iterations for nested groups */
var MAX_UNGROUP_ITERATIONS = 50;

// =========================================
// ローカライズ / Localization
// =========================================

/**
 * 実行環境の言語を判定します。
 *
 * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
 */
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
        ungroupOptions: { ja: "オプション", en: "Options" },
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
        processSymbolItem: { ja: "シンボルのリンク解除・正規化・命名", en: "Break, normalize, and name symbol" }
    }
};

/**
 * ドットパスで指定したラベルを、現在の言語で取得します。
 *
 * @param {string} labelPath - ラベルのドットパス（例 "panel.resultName"）。
 * @returns {string} 現在の言語のラベル。見つからない場合は英語、それもなければ labelPath。
 */
function getLabel(labelPath) {
    var pathParts = labelPath.split(".");
    var entry = LABELS;
    for (var i = 0; i < pathParts.length; i++) {
        if (!entry) break;
        entry = entry[pathParts[i]];
    }
    if (entry && entry[lang]) return entry[lang];
    if (entry && entry.en) return entry.en;
    return labelPath;
}

/**
 * ラベルの末尾にコロンを付けて取得します（日本語は全角、英語は半角）。
 *
 * @param {string} labelPath - ラベルのドットパス。
 * @returns {string} コロン付きのラベル。
 */
function getLabelWithColon(labelPath) {
    return getLabel(labelPath) + (lang === "ja" ? "：" : ":");
}

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

/**
 * ウィンドウに共通のレイアウトを適用します。
 *
 * @param {object} win - 対象の Window。
 * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルに共通のレイアウトを適用します。
 *
 * @param {object} panel - 対象の Panel。
 * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 横並びのグループ（ボタン列など）に共通のレイアウトを適用します。
 *
 * @param {object} group - 対象の Group。
 * @param {string} [alignment] - グループ自体の配置。省略時は "left"。
 * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignChildren = ["left", "center"];
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

(function () {

    if (app.documents.length === 0) {
        return;
    }

    var activeDocument = app.activeDocument;

    if (activeDocument.selection.length === 0) {
        return;
    }

    var ungroupSingleItemGroup = UNGROUP_SINGLE_ITEM_GROUP_DEFAULT;
    var ungroupAll = UNGROUP_ALL_DEFAULT;
    var inheritSymbolName = INHERIT_SYMBOL_NAME_DEFAULT;
    var usePrefix = USE_PREFIX_DEFAULT;
    var unlinkedItemNamePrefix = UNLINKED_ITEM_NAME_PREFIX_DEFAULT;
    var useTextContentAsName = USE_TEXT_CONTENT_AS_NAME_DEFAULT;

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オプションダイアログを表示し、選択された設定を返します。
     *
     * @returns {object} 設定オブジェクト。キャンセルされた場合は null。
     */
    function showOptionsDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var ungroupPanel = dialog.add("panel", undefined, getLabel("panel.ungroupOptions"));
        setupPanel(ungroupPanel, 6);

        var ungroupSingleItemCheckbox = ungroupPanel.add("checkbox", undefined, getLabel("checkbox.ungroupSingleItem"));
        ungroupSingleItemCheckbox.value = UNGROUP_SINGLE_ITEM_GROUP_DEFAULT;
        ungroupSingleItemCheckbox.helpTip = getLabel("tooltip.ungroupSingleItem");

        var ungroupAllCheckbox = ungroupPanel.add("checkbox", undefined, getLabel("checkbox.ungroupAll"));
        ungroupAllCheckbox.value = UNGROUP_ALL_DEFAULT;
        ungroupAllCheckbox.helpTip = getLabel("tooltip.ungroupAll");

        var resultNamePanel = dialog.add("panel", undefined, getLabel("panel.resultName"));
        setupPanel(resultNamePanel, 6);

        var inheritNameCheckbox = resultNamePanel.add("checkbox", undefined, getLabel("checkbox.inheritSymbolName"));
        inheritNameCheckbox.value = INHERIT_SYMBOL_NAME_DEFAULT;
        inheritNameCheckbox.helpTip = getLabel("tooltip.inheritSymbolName");

        var prefixRow = resultNamePanel.add("group");
        setupRow(prefixRow, "left", 6);

        var prefixCheckbox = prefixRow.add("checkbox", undefined, getLabelWithColon("checkbox.prefix"));
        prefixCheckbox.value = USE_PREFIX_DEFAULT;
        prefixCheckbox.helpTip = getLabel("tooltip.prefix");

        var prefixInput = prefixRow.add("edittext", undefined, UNLINKED_ITEM_NAME_PREFIX_DEFAULT);
        prefixInput.characters = 20;
        prefixInput.helpTip = getLabel("tooltip.prefix");

        var useTextNameCheckbox = resultNamePanel.add("checkbox", undefined, getLabel("checkbox.useTextContentAsName"));
        useTextNameCheckbox.value = USE_TEXT_CONTENT_AS_NAME_DEFAULT;
        useTextNameCheckbox.helpTip = getLabel("tooltip.useTextContentAsName");

        /**
         * チェックボックスの状態に応じて、各コントロールの有効・無効を切り替えます。
         *
         * @returns {void}
         */
        function syncEnabledStates() {
            var isFullUngroup = ungroupAllCheckbox.value;
            /* 完全解除が ON なら 1アイテム解除は無意味なので無効化 / When full ungroup is on, the single-item option is moot, so disable it */
            ungroupSingleItemCheckbox.enabled = !isFullUngroup;
            /* 完全解除が ON ならコンテナに付けた名前は破棄されるため、シンボル名系の命名を無効化 / When full ungroup is on, container names are discarded, so disable the symbol-name options */
            inheritNameCheckbox.enabled = !isFullUngroup;
            prefixCheckbox.enabled = !isFullUngroup && inheritNameCheckbox.value;
            prefixInput.enabled = prefixCheckbox.enabled && prefixCheckbox.value;
        }
        syncEnabledStates();
        ungroupSingleItemCheckbox.onClick = syncEnabledStates;
        ungroupAllCheckbox.onClick = syncEnabledStates;
        inheritNameCheckbox.onClick = syncEnabledStates;
        prefixCheckbox.onClick = syncEnabledStates;

        var buttonRow = dialog.add("group");
        setupRow(buttonRow, "right");
        buttonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering */
        dialog.onShow = function () {
            dialog.layout.layout(true);
            dialog.layout.resize();
        };

        if (dialog.show() !== 1) return null;

        return {
            ungroupSingleItem: ungroupSingleItemCheckbox.value,
            ungroupAll: ungroupAllCheckbox.value,
            inheritSymbolName: inheritNameCheckbox.value,
            usePrefix: prefixCheckbox.value,
            unlinkedPrefix: prefixInput.text,
            useTextContentAsName: useTextNameCheckbox.value
        };
    }

    // =========================================
    // 安全な実行ヘルパー / Safe execution helpers
    // =========================================

    /**
     * コンテキスト付きでエラーを $.writeln に出力します。
     *
     * @param {string} context - どの処理で起きたかを示す文言。
     * @param {object} errorObject - エラーオブジェクトまたはメッセージ。
     * @returns {void}
     */
    function logScriptError(context, errorObject) {
        $.writeln("[" + SCRIPT_NAME + " " + SCRIPT_VERSION + "] " + context + ": " + errorObject);
    }

    /**
     * 処理を実行し、失敗した場合はログを出して続行します。
     * ロック・非表示・削除済みなど、Illustrator 側の状態で失敗しうる操作に使います。
     *
     * @param {function} operation - 実行する処理。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true、失敗したら false。
     */
    function runSafely(operation, context) {
        try {
            operation();
            return true;
        } catch (e) {
            logScriptError(context, e);
            return false;
        }
    }

    /**
     * アイテムを削除します（失敗しても続行）。
     *
     * @param {object} item - 対象のアイテムまたはレイヤー。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true。
     */
    function removeItemSafely(item, context) {
        return runSafely(function () {
            item.remove();
        }, context);
    }

    /**
     * アイテムを移動します（失敗しても続行）。
     *
     * @param {object} item - 対象のアイテム。
     * @param {object} destination - 移動先。
     * @param {object} placement - ElementPlacement の値。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true。
     */
    function moveItemSafely(item, destination, placement, context) {
        return runSafely(function () {
            item.move(destination, placement);
        }, context);
    }

    /**
     * アイテムの選択状態を変更します（失敗しても続行）。
     *
     * @param {object} item - 対象のアイテム。
     * @param {boolean} isSelected - 選択するなら true。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true。
     */
    function setItemSelectedSafely(item, isSelected, context) {
        return runSafely(function () {
            item.selected = isSelected;
        }, context);
    }

    /**
     * メニューコマンドを実行します（失敗しても続行）。
     *
     * @param {string} commandName - メニューコマンド名。
     * @param {string} context - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true。
     */
    function executeMenuCommandSafely(commandName, context) {
        return runSafely(function () {
            app.executeMenuCommand(commandName);
        }, context);
    }

    // =========================================
    // 汎用ユーティリティ / Generic utilities
    // =========================================

    /**
     * ExtendScript のコレクションを通常の配列へコピーします。
     *
     * @param {object} collection - pageItems や selection などのコレクション。
     * @returns {array} コピーした配列。
     */
    function collectionToArray(collection) {
        var items = [];
        for (var i = 0; i < collection.length; i++) {
            items.push(collection[i]);
        }
        return items;
    }

    /**
     * breakLink 前後の差分から、新しく生成されたアイテムだけを抽出します。
     *
     * @param {object} currentItems - 現在のコレクションまたは配列。
     * @param {array} existingItems - 処理前に控えておいた配列。
     * @returns {array} 新しく生成されたアイテムの配列。
     */
    function collectGeneratedItems(currentItems, existingItems) {
        var generatedItems = [];
        for (var i = 0; i < currentItems.length; i++) {
            var currentItem = currentItems[i];
            var isExisting = false;
            for (var j = 0; j < existingItems.length; j++) {
                if (existingItems[j] === currentItem) {
                    isExisting = true;
                    break;
                }
            }
            if (!isExisting) generatedItems.push(currentItem);
        }
        return generatedItems;
    }

    /**
     * 指定アイテムを含む Layer を返します。
     *
     * @param {object} pageItem - 対象のアイテム。
     * @returns {object} 含まれる Layer。見つからない場合は null。
     */
    function getContainingLayer(pageItem) {
        var currentParent = pageItem.parent;
        while (currentParent && currentParent.typename !== "Layer") {
            currentParent = currentParent.parent;
        }
        return currentParent;
    }

    /**
     * 選択範囲に GroupItem が含まれるかを判定します。
     *
     * @param {object} selectedItems - selection または配列。
     * @returns {boolean} 含まれていれば true。
     */
    function containsGroupItem(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") return true;
        }
        return false;
    }

    // =========================================
    // シンボルの収集 / Symbol collection
    // =========================================

    /**
     * アイテムを再帰的にたどって SymbolItem を集めます（グループ内も対象）。
     *
     * @param {object} pageItem - 対象のアイテム。
     * @param {array} symbolItems - 収集先の配列。
     * @returns {void}
     */
    function collectSymbolItemsRecursively(pageItem, symbolItems) {
        if (!pageItem) return;

        if (pageItem.typename === "SymbolItem") {
            symbolItems.push(pageItem);
            return;
        }

        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                collectSymbolItemsRecursively(pageItem.pageItems[i], symbolItems);
            }
        }
    }

    /**
     * 選択範囲全体から SymbolItem をまとめて集めます。
     *
     * @param {object} selectedItems - selection。
     * @returns {array} SymbolItem の配列。
     */
    function collectSymbolItemsFromSelection(selectedItems) {
        var symbolItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            collectSymbolItemsRecursively(selectedItems[i], symbolItems);
        }
        return symbolItems;
    }

    // =========================================
    // 解除結果の整理 / Break-result normalization
    // =========================================

    /**
     * 指定サブレイヤーを親 Layer 直下のグループへ変換します。
     * 中身が1アイテムだけのグループは、設定に応じて解除します。
     *
     * @param {object} parentLayer - 変換先の親 Layer。
     * @param {object} generatedSubLayer - breakLink で生成されたサブレイヤー。
     * @returns {object} 変換後のグループまたは単体アイテム。変換できない場合は null。
     */
    function convertSubLayerToGroup(parentLayer, generatedSubLayer) {
        if (!parentLayer || !generatedSubLayer) return null;

        flattenSubLayersToGroups(generatedSubLayer);

        var convertedGroup = parentLayer.groupItems.add();
        for (var i = generatedSubLayer.pageItems.length - 1; i >= 0; i--) {
            moveItemSafely(generatedSubLayer.pageItems[i], convertedGroup, ElementPlacement.PLACEATBEGINNING, getLabel("log.moveStaticSubLayerItem"));
        }

        removeItemSafely(generatedSubLayer, getLabel("log.removeConvertedSubLayer"));

        if (ungroupSingleItemGroup && convertedGroup.pageItems.length === 1) {
            var singleItem = convertedGroup.pageItems[0];
            moveItemSafely(singleItem, convertedGroup, ElementPlacement.PLACEBEFORE, getLabel("log.moveSingleItemBeforeConvertedGroup"));
            removeItemSafely(convertedGroup, getLabel("log.removeSingleItemConvertedGroup"));
            return singleItem;
        }

        return convertedGroup;
    }

    /**
     * 指定 Layer 配下のサブレイヤーを、再帰的にグループへ変換してフラット化します。
     *
     * @param {object} parentLayer - 対象の Layer。
     * @returns {void}
     */
    function flattenSubLayersToGroups(parentLayer) {
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            convertSubLayerToGroup(parentLayer, parentLayer.layers[i]);
        }
    }

    /**
     * 複数のアイテムはグループにまとめ、1つだけならそのまま返します。
     *
     * @param {object} targetLayer - グループを作成する Layer。
     * @param {array} items - 対象のアイテム配列。
     * @param {string} logLabelPath - 移動失敗時にログへ出すラベルのドットパス。
     * @returns {object} 作成したグループまたは単体アイテム。アイテムがなければ null。
     */
    function groupItemsOrReturnSingle(targetLayer, items, logLabelPath) {
        if (items.length === 0) return null;
        if (items.length === 1) return items[0];

        var resultGroup = targetLayer.groupItems.add();
        for (var i = 0; i < items.length; i++) {
            moveItemSafely(items[i], resultGroup, ElementPlacement.PLACEATEND, getLabel(logLabelPath));
        }
        return resultGroup;
    }

    /**
     * static シンボル（サブレイヤー生成型）の解除結果を、グループまたは単体に整理します。
     *
     * @param {object} targetLayer - 解除対象があった Layer。
     * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
     * @param {array} generatedSubLayers - breakLink で生成されたサブレイヤーの配列。
     * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
     */
    function organizeStaticBreakResult(targetLayer, itemsBeforeBreak, generatedSubLayers) {
        if (!targetLayer) return null;

        for (var i = 0; i < generatedSubLayers.length; i++) {
            convertSubLayerToGroup(targetLayer, generatedSubLayers[i]);
        }

        var generatedItems = collectGeneratedItems(targetLayer.pageItems, itemsBeforeBreak);
        return groupItemsOrReturnSingle(targetLayer, generatedItems, "log.moveStaticBreakItemIntoResultGroup");
    }

    /**
     * dynamic シンボル（ページアイテム生成型）の解除結果を、
     * ungroup で平坦化したうえでグループまたは単体に整理します。
     *
     * @param {object} documentObject - 対象ドキュメント。
     * @param {object} targetLayer - 解除対象があった Layer。
     * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
     * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
     */
    function organizeDynamicBreakResult(documentObject, targetLayer, itemsBeforeBreak) {
        if (!targetLayer) return null;

        /* breakLink 直後の selection に依存せず、Layer 配下の差分で新規アイテムを特定 / Detect generated items by Layer reference differences instead of relying on the selection after breakLink */
        var generatedItems = collectGeneratedItems(targetLayer.pageItems, itemsBeforeBreak);
        if (generatedItems.length === 0) return null;

        documentObject.selection = null;
        for (var i = 0; i < generatedItems.length; i++) {
            setItemSelectedSafely(generatedItems[i], true, getLabel("log.selectDynamicBreakItem"));
        }

        /* ネストが深い場合に備え、ungroup の反復回数に上限を設ける / Cap the ungroup iterations to guard against deeply nested groups */
        var ungroupCount = 0;
        while (containsGroupItem(documentObject.selection) && ungroupCount < MAX_UNGROUP_ITERATIONS) {
            if (!executeMenuCommandSafely("ungroup", getLabel("log.ungroupNestedDynamicBreakGroup"))) break;
            ungroupCount++;
        }

        /* selection は移動中に変化するため、事前に配列へ退避 / Snapshot the selection into an array since it changes during moves */
        var brokenItems = collectionToArray(documentObject.selection);
        if (brokenItems.length === 0) return null;

        var resultLayer = getContainingLayer(brokenItems[0]) || targetLayer;
        return groupItemsOrReturnSingle(resultLayer, brokenItems, "log.moveDynamicBreakItemIntoResultGroup");
    }

    /**
     * 生成物の種類から、解除結果の型を判定します。
     *
     * @param {array} generatedSubLayers - 生成されたサブレイヤーの配列。
     * @param {array} generatedPageItems - 生成されたページアイテムの配列。
     * @returns {string} "static" / "dynamic" / "mixed" / "none" のいずれか。
     */
    function classifyBreakResult(generatedSubLayers, generatedPageItems) {
        var hasSubLayers = generatedSubLayers.length > 0;
        var hasPageItems = generatedPageItems.length > 0;

        if (hasSubLayers && !hasPageItems) return "static";
        if (!hasSubLayers && hasPageItems) return "dynamic";
        if (hasSubLayers && hasPageItems) return "mixed";
        return "none";
    }

    /**
     * 判定結果に応じて、static / dynamic の整理処理へ振り分けます。
     *
     * @param {object} documentObject - 対象ドキュメント。
     * @param {object} targetLayer - 解除対象があった Layer。
     * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
     * @param {array} generatedSubLayers - 生成されたサブレイヤーの配列。
     * @param {array} generatedPageItems - 生成されたページアイテムの配列。
     * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
     */
    function normalizeBreakResult(documentObject, targetLayer, itemsBeforeBreak, generatedSubLayers, generatedPageItems) {
        var breakResultType = classifyBreakResult(generatedSubLayers, generatedPageItems);

        if (breakResultType === "dynamic") {
            return organizeDynamicBreakResult(documentObject, targetLayer, itemsBeforeBreak);
        }

        if (breakResultType === "mixed") {
            logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.mixedBreakLinkResult"));
        }

        if (breakResultType === "none") {
            logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.noBreakLinkResult"));
            return null;
        }

        return organizeStaticBreakResult(targetLayer, itemsBeforeBreak, generatedSubLayers);
    }

    // =========================================
    // 解除後の命名 / Result naming
    // =========================================

    /**
     * 設定に応じて、接頭辞を付けた解除後の名前を組み立てます。
     *
     * @param {string} symbolName - 元のシンボル名。
     * @returns {string} 解除後のアイテム名。
     */
    function buildUnlinkedItemName(symbolName) {
        return (usePrefix ? unlinkedItemNamePrefix : "") + symbolName;
    }

    /**
     * 単一 TextFrame の文字列を取得します（改行は空白に置換）。
     *
     * @param {object} item - 対象のアイテム。
     * @returns {string} テキストの内容。対象外または空文字の場合は null。
     */
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
        return contents.length === 0 ? null : contents;
    }

    /**
     * 解除結果アイテムに名前を付けます（単一テキストの内容を元シンボル名より優先）。
     *
     * @param {object} resultItem - 解除結果のアイテム。
     * @param {string} symbolName - 元のシンボル名。
     * @returns {void}
     */
    function applyResultName(resultItem, symbolName) {
        if (!resultItem) return;

        if (useTextContentAsName) {
            var textContent = getSingleTextFrameContent(resultItem);
            if (textContent !== null) {
                resultItem.name = textContent;
                return;
            }
        }

        if (inheritSymbolName) {
            resultItem.name = buildUnlinkedItemName(symbolName);
        }
    }

    // =========================================
    // シンボル1つ分の処理 / Per-symbol processing
    // =========================================

    /**
     * breakLink 前の Layer 直下のアイテムとサブレイヤーを控えておきます。
     *
     * @param {object} targetLayer - 対象の Layer。
     * @returns {object} pageItems / subLayers を持つオブジェクト。
     */
    function snapshotLayerContents(targetLayer) {
        return {
            pageItems: targetLayer ? collectionToArray(targetLayer.pageItems) : [],
            subLayers: targetLayer ? collectionToArray(targetLayer.layers) : []
        };
    }

    /**
     * シンボルのリンクを解除し、生成物を整理して1つの結果アイテムにまとめます。
     *
     * @param {object} symbolItem - 対象の SymbolItem。
     * @param {object} documentObject - 対象ドキュメント。
     * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
     */
    function breakAndNormalizeSymbol(symbolItem, documentObject) {
        var targetLayer = getContainingLayer(symbolItem);
        var beforeBreak = snapshotLayerContents(targetLayer);

        symbolItem.breakLink();

        var generatedSubLayers = targetLayer ? collectGeneratedItems(targetLayer.layers, beforeBreak.subLayers) : [];
        var generatedPageItems = targetLayer ? collectGeneratedItems(targetLayer.pageItems, beforeBreak.pageItems) : [];

        return normalizeBreakResult(documentObject, targetLayer, beforeBreak.pageItems, generatedSubLayers, generatedPageItems);
    }

    /**
     * 解除結果グループをネストごと完全に解除し、展開後のアイテム配列を返します。
     *
     * @param {object} documentObject - 対象ドキュメント。
     * @param {object} resultItem - 解除結果のアイテム。
     * @returns {array} 展開後のアイテム配列。
     */
    function ungroupResultCompletely(documentObject, resultItem) {
        if (!resultItem) return [];
        if (resultItem.typename !== "GroupItem") return [resultItem];

        documentObject.selection = null;
        if (!setItemSelectedSafely(resultItem, true, getLabel("log.ungroupAllResultGroup"))) return [resultItem];

        /* ungroupAll はネストごと一括で解除する / ungroupAll dissolves the group and all nested groups at once */
        executeMenuCommandSafely("ungroupAll", getLabel("log.ungroupAllResultGroup"));

        /* 解除直後の選択が展開後アイテム / The selection right after ungroupAll holds the loose items */
        return collectionToArray(documentObject.selection);
    }

    /**
     * SymbolItem 1つ分の処理（解除・整理・命名・完全解除）をまとめます。
     *
     * @param {object} symbolItem - 対象の SymbolItem。
     * @param {object} documentObject - 対象ドキュメント。
     * @returns {array} 選択状態として残す結果アイテムの配列。
     */
    function processSymbolItem(symbolItem, documentObject) {
        if (symbolItem.locked || symbolItem.hidden) return [];

        /* 外側で初期選択を解除済みの前提で、この SymbolItem だけを処理対象として選択 / Select only this SymbolItem, assuming the outer flow already cleared the initial selection */
        if (!setItemSelectedSafely(symbolItem, true, getLabel("log.selectTargetSymbolItem"))) return [];

        var symbolName = symbolItem.symbol.name;
        var resultItem = breakAndNormalizeSymbol(symbolItem, documentObject);

        applyResultName(resultItem, symbolName);

        /* 完全解除が ON なら、命名後に結果グループをネストごと解除して展開後アイテムを返す / When full ungroup is on, dissolve the result group after naming and return the loose items */
        if (ungroupAll) return ungroupResultCompletely(documentObject, resultItem);

        return resultItem ? [resultItem] : [];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    var symbolItems = collectSymbolItemsFromSelection(activeDocument.selection);

    if (symbolItems.length === 0) {
        return;
    }

    if (SHOW_OPTIONS_DIALOG) {
        var dialogSettings = showOptionsDialog();
        if (!dialogSettings) {
            return;
        }
        ungroupSingleItemGroup = dialogSettings.ungroupSingleItem;
        ungroupAll = dialogSettings.ungroupAll;
        inheritSymbolName = dialogSettings.inheritSymbolName;
        usePrefix = dialogSettings.usePrefix;
        unlinkedItemNamePrefix = dialogSettings.unlinkedPrefix;
        useTextContentAsName = dialogSettings.useTextContentAsName;
    }

    /* 処理後の選択方針 / Post-processing selection policy */
    /* 元の選択は復元せず、解除後に生成されたアイテムを選択状態として残す / Do not restore the original selection; keep the generated unlinked items selected */
    /* 初期選択はここで一度だけ解除し、生成物を配列に集めて末尾でまとめて選択する / Clear the initial selection once here, collect generated items, and select them all at the end */
    activeDocument.selection = null;

    var generatedResultItems = [];
    for (var i = 0; i < symbolItems.length; i++) {
        /* 1つのシンボルで失敗しても、残りのシンボルの処理は続行する / Keep processing the remaining symbols even if one of them fails */
        try {
            var producedItems = processSymbolItem(symbolItems[i], activeDocument);
            for (var j = 0; j < producedItems.length; j++) {
                generatedResultItems.push(producedItems[j]);
            }
        } catch (e) {
            logScriptError(getLabel("log.processSymbolItem"), e);
        }
    }

    /* 全シンボルの生成物を最終的に選択状態へ（static / dynamic / 完全解除で共通）/ Select every symbol's generated items at the end (uniform for static / dynamic / full-ungroup) */
    activeDocument.selection = null;
    for (var k = 0; k < generatedResultItems.length; k++) {
        setItemSelectedSafely(generatedResultItems[k], true, getLabel("log.selectGeneratedResultItem"));
    }

})();
