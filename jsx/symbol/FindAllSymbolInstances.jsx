#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * Find All Instances of Selected Symbol Instances
 * 更新日: 2026-05-09
 *
 * 概要:
 *  - 選択中の SymbolItem（GroupItem 内にネストされたものも含む）を再帰的に集めます。
 *  - 集めたインスタンスのうち、シンボル定義ごとに「最初に出現した1つ」を代表として選び、
 *    代表ごとに app.executeMenuCommand('Find Symbol Instance menu item');
 *    を実行して、そのシンボルの全インスタンスを得ます。
 *  - 各実行結果を重複なくマージし、最終的に「選択していたシンボル群すべての全インスタンス」を
 *    まとめて再選択します（複数のシンボルが混在している場合にも対応）。
 *  - 選択に SymbolItem が1つも含まれない場合は、フォールバックとして
 *      app.executeMenuCommand('Find Appearance menu item');
 *    を実行します（同一アピアランスのオブジェクトを検索）。
 *  - ロック等で再選択できないアイテムは黙ってスキップします。
 *  - 何も集約できなかった場合のみアラートを表示します。
 */

// =========================================
// バージョン情報 / Version
// =========================================

var SCRIPT_VERSION = "v1.1.0";

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    if (doc.selection.length === 0) return;

    /* 選択から SymbolItem を再帰抽出 / Collect SymbolItems from selection (recursive) */
    var selectedSymbolItems = collectSymbolItemsFromSelection(doc.selection);

    /* シンボルが含まれない場合は Find Appearance を実行 / Fall back to Find Appearance when no SymbolItem is selected */
    if (selectedSymbolItems.length === 0) {
        runFindAppearance();
        return;
    }

    /* シンボル定義ごとの代表インスタンスを抽出 / Pick one representative per symbol definition */
    var representativeInstances = pickRepresentativePerSymbol(selectedSymbolItems);

    /* 代表ごとに Find Symbol Instance を実行して結果を集約 / Run Find Symbol Instance per representative and collect results */
    var allMatchedInstances = collectAllSiblingInstances(doc, representativeInstances);

    /* 集約した全インスタンスをまとめて選択 / Re-select all collected instances */
    applySelection(doc, allMatchedInstances);

    if (allMatchedInstances.length === 0) {
        alert("同じインスタンスの選択に失敗しました（コマンドが無効/選択対象が不正など）。");
    }

    // =========================================
    // ヘルパー: 選択 → SymbolItem 抽出 / Helpers: Selection -> SymbolItem
    // =========================================

    /* 選択配列から SymbolItem を全て集める / Collect every SymbolItem inside the selection array */
    function collectSymbolItemsFromSelection(selection) {
        var result = [];
        for (var i = 0; i < selection.length; i++) {
            collectSymbolItemsDeep(selection[i], result);
        }
        return result;
    }

    /* 単一アイテムを再帰的に走査して SymbolItem を resultList に追加 / Walk a page item recursively and push SymbolItems into resultList */
    function collectSymbolItemsDeep(pageItem, resultList) {
        if (!pageItem) return;
        if (pageItem.typename === "SymbolItem") {
            resultList.push(pageItem);
            return;
        }
        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                collectSymbolItemsDeep(pageItem.pageItems[i], resultList);
            }
        }
    }

    // =========================================
    // ヘルパー: シンボル代表の抽出 / Helpers: Pick representative
    // =========================================

    /* シンボル定義ごとに最初に出現したインスタンスを1つだけ代表として返す / Return one representative instance per unique symbol definition */
    function pickRepresentativePerSymbol(symbolItems) {
        var seenSymbols = [];
        var representatives = [];
        for (var i = 0; i < symbolItems.length; i++) {
            var symbolDef = symbolItems[i].symbol;
            if (!symbolDef || arrayContains(seenSymbols, symbolDef)) continue;
            seenSymbols.push(symbolDef);
            representatives.push(symbolItems[i]);
        }
        return representatives;
    }

    // =========================================
    // ヘルパー: 同一シンボルの全インスタンス収集 / Helpers: Collect sibling instances
    // =========================================

    /* 代表ごとに Find Symbol Instance を実行し、選択結果を重複なく集約 / For each representative run Find Symbol Instance and merge selections without duplicates */
    function collectAllSiblingInstances(doc, representatives) {
        var collected = [];
        for (var i = 0; i < representatives.length; i++) {
            deselectAll(doc);
            representatives[i].selected = true;
            try {
                app.executeMenuCommand('Find Symbol Instance menu item');
            } catch (findErr) {
                continue;
            }
            var currentSelection = doc.selection;
            for (var j = 0; j < currentSelection.length; j++) {
                var item = currentSelection[j];
                if (item && !arrayContains(collected, item)) {
                    collected.push(item);
                }
            }
        }
        return collected;
    }

    // =========================================
    // ヘルパー: 選択操作・メニュー実行 / Helpers: Selection ops & menu commands
    // =========================================

    /* 渡された配列をまとめて選択（ロック等は黙ってスキップ） / Select all items in the array (skip silently when locked or unselectable) */
    function applySelection(doc, items) {
        deselectAll(doc);
        for (var i = 0; i < items.length; i++) {
            try { items[i].selected = true; } catch (selectErr) {}
        }
    }

    /* シンボル不在時の代替コマンド: Find Appearance を実行 / Fallback command when no SymbolItem is selected: run Find Appearance */
    function runFindAppearance() {
        try {
            app.executeMenuCommand('Find Appearance menu item');
        } catch (findApprErr) {}
    }

    /* 全選択を解除 / Clear the current selection */
    function deselectAll(targetDoc) {
        try { targetDoc.selection = null; } catch (deselectErr) {}
    }

    // =========================================
    // ヘルパー: 汎用ユーティリティ / Helpers: Generic utilities
    // =========================================

    /* 配列内に同一参照が存在するか判定 / Check whether a list already contains the target reference */
    function arrayContains(list, target) {
        for (var i = 0; i < list.length; i++) {
            if (list[i] === target) return true;
        }
        return false;
    }
})();
