#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SelectSameAppearanceMulti
Name=複数のオブジェクトを基準に「共通 > アピアランス」を実行
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.0
Release-Date=2026-09-06
Description-BEGIN
選択したオブジェクトそれぞれに「共通 > アピアランス」（シンボルインスタンスは「共通 > シンボルインスタンス」）を適用し、見つかったオブジェクトをまとめて選択し直します。基準に〈塗り〉なし・〈線〉なしのパスがあれば、〈塗り〉なし・〈線〉なしのパスもあわせて選択します。標準機能では基準にできるオブジェクトが1つだけという制限を回避できます。

Applies "Select > Same > Appearance" (or "Same > Symbol Instance" for symbol instances) to each selected object and reselects every object found. If a reference path has no fill or no stroke, every path with no fill or no stroke is selected as well. This works around the built-in limitation of using only one object as the reference.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトそれぞれに「共通 > アピアランス」（シンボルインスタンスは「共通 > シンボルインスタンス」）を適用し、見つかったオブジェクトをまとめて選択し直します。基準に〈塗り〉なし・〈線〉なしのパスがあれば、〈塗り〉なし・〈線〉なしのパスもあわせて選択します。標準機能では基準にできるオブジェクトが1つだけという制限を回避できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md

### Overview

Applies "Select > Same > Appearance" (or "Same > Symbol Instance" for symbol instances) to each selected object and reselects every object found. If a reference path has no fill or no stroke, every path with no fill or no stroke is selected as well. This works around the built-in limitation of using only one object as the reference.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectSameAppearanceMulti";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-21";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "複数のオブジェクトを基準に「共通 > アピアランス」を実行"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectSameAppearanceMulti.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ローカライズ（再利用パーツ） / Localization (reusable)

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* UIラベル定義 / UI label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select at least one object." }
        }
    };

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    /* doc.selection は参照するたびに新しい配列を返すため、そのまま基準リストとして保持できる */
    var referencePageItems = doc.selection;

    if (!referencePageItems || referencePageItems.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /**
     * 指定したページアイテムを基準に「共通 > アピアランス」を実行し、選択されたページアイテムを返す。
     * シンボルインスタンスは「共通 > シンボルインスタンス」で同じシンボルのインスタンスを選択する
     * @param {PageItem} referencePageItem - 基準にするページアイテム
     * @returns {PageItem[]} 一致したページアイテムの配列
     */
    function findPageItemsWithSameAppearance(referencePageItem) {
        doc.selection = null;
        referencePageItem.selected = true;
        /* 選択状態を確定させてからメニューコマンドを実行する */
        app.redraw();
        if (referencePageItem.typename === "SymbolItem") {
            app.executeMenuCommand("Find Symbol Instance menu item");
        } else {
            app.executeMenuCommand("Find Appearance menu item");
        }
        return doc.selection;
    }

    /**
     * パスまたは複合パスの塗り・線の有無を返す。複合パスは先頭のパスの値を使う
     * @param {PageItem} pageItem - 調べるページアイテム
     * @returns {{filled: boolean, stroked: boolean}|null} 塗り・線の有無。パス・複合パス以外は null
     */
    function getFillStrokeState(pageItem) {
        var pathItem = pageItem;
        if (pageItem.typename === "CompoundPathItem") {
            if (pageItem.pathItems.length === 0) return null;
            pathItem = pageItem.pathItems[0];
        } else if (pageItem.typename !== "PathItem") {
            return null;
        }
        return { filled: pathItem.filled, stroked: pathItem.stroked };
    }

    /**
     * ページアイテムが選択できる状態か（自身と親のグループ・レイヤーがロック・非表示でないか）を返す
     * @param {PageItem} pageItem - 調べるページアイテム
     * @returns {boolean} 選択できれば true
     */
    function isPageItemSelectable(pageItem) {
        if (pageItem.locked || pageItem.hidden) return false;
        var parentItem = pageItem.parent;
        while (parentItem && parentItem.typename !== "Document") {
            if (parentItem.typename === "Layer") {
                if (parentItem.locked || !parentItem.visible) return false;
            } else if (parentItem.locked || parentItem.hidden) {
                return false;
            }
            parentItem = parentItem.parent;
        }
        return true;
    }

    /**
     * 〈塗り〉なし、または〈線〉なしのパス・複合パスをドキュメントから集める。
     * ガイド・クリッピングパス・複合パスの中のパス・ロックや非表示のものは除く
     * @param {boolean} collectNoFill - 〈塗り〉なしを集めるか
     * @param {boolean} collectNoStroke - 〈線〉なしを集めるか
     * @returns {PageItem[]} 該当したページアイテムの配列
     */
    function findNoFillOrNoStrokePageItems(collectNoFill, collectNoStroke) {
        var foundPageItems = [];
        var candidateLists = [doc.pathItems, doc.compoundPathItems];
        for (var listIndex = 0; listIndex < candidateLists.length; listIndex++) {
            var candidateList = candidateLists[listIndex];
            for (var itemIndex = 0; itemIndex < candidateList.length; itemIndex++) {
                var candidateItem = candidateList[itemIndex];
                if (candidateItem.typename === "PathItem") {
                    if (candidateItem.guides || candidateItem.clipping) continue;
                    if (candidateItem.parent.typename === "CompoundPathItem") continue;
                }
                var fillStrokeState = getFillStrokeState(candidateItem);
                if (!fillStrokeState) continue;
                var isMatch = (collectNoFill && !fillStrokeState.filled) || (collectNoStroke && !fillStrokeState.stroked);
                if (!isMatch || !isPageItemSelectable(candidateItem)) continue;
                foundPageItems.push(candidateItem);
            }
        }
        return foundPageItems;
    }

    /* 重複したまま集めてよい。選択し直す際に同じページアイテムを複数回指定しても結果は変わらない */
    var pageItemsToSelect = [];
    /* 基準に〈塗り〉なし・〈線〉なしのパスがあるか / Whether any reference path has no fill or no stroke */
    var hasNoFillReference = false;
    var hasNoStrokeReference = false;
    /* 検索済みのシンボル名。同じシンボルのインスタンスは結果が同じなので1回だけ検索する / Symbol names already searched */
    var searchedSymbolNames = {};
    for (var i = 0; i < referencePageItems.length; i++) {
        var referenceFillStroke = getFillStrokeState(referencePageItems[i]);
        if (referenceFillStroke) {
            if (!referenceFillStroke.filled) hasNoFillReference = true;
            if (!referenceFillStroke.stroked) hasNoStrokeReference = true;
        }
        if (referencePageItems[i].typename === "SymbolItem") {
            /* シンボル名はドキュメント内で一意 / Symbol names are unique within a document */
            var symbolNameKey = "#" + referencePageItems[i].symbol.name;
            if (searchedSymbolNames[symbolNameKey]) continue;
            searchedSymbolNames[symbolNameKey] = true;
        }
        var samePageItems = findPageItemsWithSameAppearance(referencePageItems[i]);
        for (var j = 0; j < samePageItems.length; j++) {
            pageItemsToSelect.push(samePageItems[j]);
        }
    }

    /* 〈塗り〉なし・〈線〉なしの基準があれば、該当するパスをすべて加える / Add every path with no fill or no stroke */
    if (hasNoFillReference || hasNoStrokeReference) {
        var noFillStrokePageItems = findNoFillOrNoStrokePageItems(hasNoFillReference, hasNoStrokeReference);
        for (var m = 0; m < noFillStrokePageItems.length; m++) {
            pageItemsToSelect.push(noFillStrokePageItems[m]);
        }
    }

    doc.selection = null;
    for (var k = 0; k < pageItemsToSelect.length; k++) {
        pageItemsToSelect[k].selected = true;
    }
    app.redraw();
})();
