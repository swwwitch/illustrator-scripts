#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のシンボルインスタンスと同じシンボルのインスタンスを、ドキュメント全体から探してまとめて選択し直します。
グループ内のインスタンスや複数シンボルの混在にも対応し、シンボルを含まない選択では同じアピアランスのオブジェクトを選択します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FindAllSymbolInstances.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n140952ad5011

### Overview

Finds every instance of the same symbols as the selected symbol instances across the document and reselects them all.
Instances inside groups and mixed symbols are handled; when no symbol instance is selected, objects with the same appearance are selected instead.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FindAllSymbolInstances.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FindAllSymbolInstances";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-15";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FindAllSymbolInstances.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FindAllSymbolInstances.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n140952ad5011"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の UI 言語 / Current UI language */
    var uiLang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            selectFailed: { ja: '同じシンボルのインスタンスを選択できませんでした。', en: 'Could not select instances of the same symbol.' }
        }
    };

    /**
     * ラベル（ja/en）を現在の UI 言語の文字列にする
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || '';
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) return;

    var activeDoc = app.activeDocument;
    if (activeDoc.selection.length === 0) return;

    /* 選択から SymbolItem を再帰抽出 / Collect SymbolItems from selection (recursive) */
    var selectedSymbolItems = collectSymbolItems(activeDoc.selection, []);

    /* シンボルが含まれない場合は Find Appearance を実行 / Fall back to Find Appearance when no SymbolItem is selected */
    if (selectedSymbolItems.length === 0) {
        app.executeMenuCommand('Find Appearance menu item');
        return;
    }

    /* シンボル定義ごとの代表インスタンスを抽出 / Pick one representative per symbol definition */
    var representativeInstances = pickOneInstancePerSymbol(selectedSymbolItems);

    /* 代表ごとに Find Symbol Instance を実行して結果を集約 / Run Find Symbol Instance per representative and collect results */
    var allMatchedInstances = collectInstancesOfSameSymbols(representativeInstances);

    /* 集約した全インスタンスをまとめて選択 / Re-select all collected instances */
    reselectItems(allMatchedInstances);

    if (allMatchedInstances.length === 0) {
        alert(getLabel(LABELS.alert.selectFailed));
    }

    // =========================================
    // ヘルパー: シンボルインスタンスの収集 / Helpers: Collect symbol instances
    // =========================================

    /**
     * アイテムの配列を再帰的に走査し、グループ内も含めて SymbolItem を集める
     * @param {PageItem[]} pageItems - 走査するアイテム（選択範囲やグループの pageItems）
     * @param {SymbolItem[]} foundSymbolItems - 見つかった SymbolItem の追加先
     * @returns {SymbolItem[]} foundSymbolItems と同じ配列
     */
    function collectSymbolItems(pageItems, foundSymbolItems) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!pageItem) continue;
            if (pageItem.typename === "SymbolItem") {
                foundSymbolItems.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectSymbolItems(pageItem.pageItems, foundSymbolItems);
            }
        }
        return foundSymbolItems;
    }

    /**
     * シンボル定義ごとに、最初に出現したインスタンスを1つだけ代表として返す
     * @param {SymbolItem[]} symbolItems - 選択から集めた SymbolItem
     * @returns {SymbolItem[]} シンボル定義が重複しない代表インスタンス
     */
    function pickOneInstancePerSymbol(symbolItems) {
        var seenSymbolNames = {};
        var pickedInstances = [];
        for (var i = 0; i < symbolItems.length; i++) {
            /* シンボル名はドキュメント内で一意 / Symbol names are unique within a document */
            var symbolNameKey = "#" + symbolItems[i].symbol.name;
            if (seenSymbolNames[symbolNameKey]) continue;
            seenSymbolNames[symbolNameKey] = true;
            pickedInstances.push(symbolItems[i]);
        }
        return pickedInstances;
    }

    /**
     * 代表ごとに Find Symbol Instance を実行し、選択されたインスタンスを集約する
     * （代表のシンボル定義はすべて異なるため、結果は重複しない）
     * @param {SymbolItem[]} pickedInstances - シンボル定義ごとの代表インスタンス
     * @returns {SymbolItem[]} 同じシンボルのインスタンスすべて
     */
    function collectInstancesOfSameSymbols(pickedInstances) {
        var matchedInstances = [];
        for (var i = 0; i < pickedInstances.length; i++) {
            activeDoc.selection = null;
            pickedInstances[i].selected = true;
            app.executeMenuCommand('Find Symbol Instance menu item');

            var foundInstances = activeDoc.selection;
            for (var j = 0; j < foundInstances.length; j++) {
                matchedInstances.push(foundInstances[j]);
            }
        }
        return matchedInstances;
    }

    // =========================================
    // ヘルパー: 選択操作 / Helpers: Selection
    // =========================================

    /**
     * 選択を解除してから、渡されたアイテムをまとめて選択する（選択できないものは黙ってスキップ）
     * @param {PageItem[]} itemsToSelect - 選択するアイテム
     * @returns {void}
     */
    function reselectItems(itemsToSelect) {
        activeDoc.selection = null;
        for (var i = 0; i < itemsToSelect.length; i++) {
            try { itemsToSelect[i].selected = true; } catch (err) {}
        }
    }
})();
