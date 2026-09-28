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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            selectFailed: { ja: '同じシンボルのインスタンスを選択できませんでした。', en: 'Could not select instances of the same symbol.' }
        }
    };

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
