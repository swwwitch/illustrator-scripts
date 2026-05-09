#target illustrator

/*
 * Find All Instances of Selected Symbol Instances
 * 更新日: 2025-12-15
 *
 * 概要:
 *  - 複数のシンボルインスタンス(SymbolItem)を選択している状態から、
 *    それらが属するシンボル定義ごとに
 *      app.executeMenuCommand('Find Symbol Instance menu item');
 *    を1つずつ実行し、最終的に「選択していたシンボルインスタンスと同じシンボルの全インスタンス」を
 *    まとめて選択します。
 */

function main() {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;

    if (!doc.selection || doc.selection.length === 0) {
        return;
    }

    // 選択から SymbolItem を再帰的に抽出（GroupItem 対応）
    var selectedSymbolItems = [];
    for (var i = 0; i < doc.selection.length; i++) {
        collectSymbolItemsDeep(doc.selection[i], selectedSymbolItems);
    }

    if (selectedSymbolItems.length === 0) {
        return;
    }

    // 選択中の「シンボル定義」を重複なしで集める
    var symbols = [];
    for (var s = 0; s < selectedSymbolItems.length; s++) {
        var sym = selectedSymbolItems[s].symbol; // Symbol
        if (sym && !arrayContains(symbols, sym)) {
            symbols.push(sym);
        }
    }

    // 結果（全インスタンス）を重複なしで集める
    var resultItems = [];

    // いったん全選択解除
    deselectAll(doc);

    // シンボル定義ごとに「Find Symbol Instance」→選択結果を回収
    for (var k = 0; k < symbols.length; k++) {
        var targetSymbol = symbols[k];

        // 代表インスタンス（最初に選択していたものの中から探す）
        var representative = findRepresentative(selectedSymbolItems, targetSymbol);
        if (!representative) continue;

        deselectAll(doc);
        representative.selected = true;

        try {
            // これが「同じシンボルの全インスタンスを選択」に相当
            app.executeMenuCommand('Find Symbol Instance menu item');
        } catch (e) {
            // コマンドが無効・名称違い等
            // ここで止めず、次のシンボルへ
            continue;
        }

        // 実行後の選択を回収
        var curSel = doc.selection;
        for (var j = 0; j < curSel.length; j++) {
            var item = curSel[j];
            if (item && !arrayContains(resultItems, item)) {
                resultItems.push(item);
            }
        }
    }

    // 最終的にまとめて選択
    deselectAll(doc);
    for (var r = 0; r < resultItems.length; r++) {
        try {
            resultItems[r].selected = true;
        } catch (e2) {
            // ロック等で選択できない場合はスキップ
        }
    }

    if (resultItems.length === 0) {
        alert("同じインスタンスの選択に失敗しました（コマンドが無効/選択対象が不正など）。");
    }

    // ---- helpers ----

    function collectSymbolItemsDeep(item, result) {
        if (!item) return;

        if (item.typename === "SymbolItem") {
            if (!arrayContains(result, item)) {
                result.push(item);
            }
            return;
        }

        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectSymbolItemsDeep(item.pageItems[i], result);
            }
        }
    }

    function arrayContains(arr, obj) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === obj) return true;
        }
        return false;
    }

    function deselectAll(doc) {
        try {
            doc.selection = null;
        } catch (e) {}
    }

    function findRepresentative(symbolItems, symbolDef) {
        for (var i = 0; i < symbolItems.length; i++) {
            if (symbolItems[i] && symbolItems[i].symbol === symbolDef) {
                return symbolItems[i];
            }
        }
        return null;
    }

}
main();