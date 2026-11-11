#target illustrator
#targetengine "SelectSameAppearanceMulti"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SelectSameAppearanceMulti
Name=複数のオブジェクトを基準に「共通 > アピアランス」を実行
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.1
Release-Date=2026-09-06
Description-BEGIN
選択したオブジェクトそれぞれに「共通 > アピアランス」（シンボルインスタンスは「共通 > シンボルインスタンス」）を適用し、見つかったオブジェクトをまとめて選択し直します。基準に〈塗り〉も〈線〉もないパスがあれば、〈塗り〉も〈線〉もないパスもあわせて選択します。標準機能では基準にできるオブジェクトが1つだけという制限を回避できます。

Applies "Select > Same > Appearance" (or "Same > Symbol Instance" for symbol instances) to each selected object and reselects every object found. If a reference path has neither fill nor stroke, every path with neither fill nor stroke is selected as well. This works around the built-in limitation of using only one object as the reference.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトそれぞれに「共通 > アピアランス」（シンボルインスタンスは「共通 > シンボルインスタンス」）を適用し、見つかったオブジェクトをまとめて選択し直します。基準に〈塗り〉も〈線〉もないパスがあれば、〈塗り〉も〈線〉もないパスもあわせて選択します。標準機能では基準にできるオブジェクトが1つだけという制限を回避できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n61c859d35d86

### Overview

Applies "Select > Same > Appearance" (or "Same > Symbol Instance" for symbol instances) to each selected object and reselects every object found. If a reference path has neither fill nor stroke, every path with neither fill nor stroke is selected as well. This works around the built-in limitation of using only one object as the reference.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectSameAppearanceMulti";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-12";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "複数のオブジェクトを基準に「共通 > アピアランス」を実行"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n61c859d35d86"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectSameAppearanceMulti.jsx"; /* 最新版のソース / latest source */

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
     * @param {boolean} isSoleSelection - 基準だけがすでに選択されているか（true なら選択を付け替えない）
     * @returns {PageItem[]} 一致したページアイテムの配列
     */
    function findPageItemsWithSameAppearance(referencePageItem, isSoleSelection) {
        if (!isSoleSelection) {
            doc.selection = null;
            referencePageItem.selected = true;
            /* 選択状態を確定させてからメニューコマンドを実行する */
            app.redraw();
        }
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
     * 〈塗り〉も〈線〉もないパス・複合パスをドキュメントから集める。
     * ガイド・クリッピングパス・複合パスの中のパス・ロックや非表示のものは除く
     * @returns {PageItem[]} 該当したページアイテムの配列
     */
    function findNoFillNoStrokePageItems() {
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
                if (fillStrokeState.filled || fillStrokeState.stroked) continue;
                if (!isPageItemSelectable(candidateItem)) continue;
                foundPageItems.push(candidateItem);
            }
        }
        return foundPageItems;
    }

    /**
     * 2つの基準で使うメニューコマンドが同じか（シンボルインスタンスどうし、またはそれ以外どうし）を返す
     * @param {PageItem} pageItemA - 基準A
     * @param {PageItem} pageItemB - 基準B
     * @returns {boolean} 同じコマンドで検索するなら true
     */
    function usesSameFindCommand(pageItemA, pageItemB) {
        return (pageItemA.typename === "SymbolItem") === (pageItemB.typename === "SymbolItem");
    }

    /* 検索ごとの結果。検索済みの基準と同じものは検索しないので、結果どうしは重ならない
       Results per search. References already found are skipped, so the results do not overlap */
    var foundGroups = [];
    /* 前の検索で見つかった基準（番号 → true）/ References found by an earlier search (index -> true) */
    var isCoveredReference = {};
    /* 基準に〈塗り〉も〈線〉もないパスがあるか / Whether any reference path has neither fill nor stroke */
    var hasNoFillNoStrokeReference = false;
    for (var i = 0; i < referencePageItems.length; i++) {
        var referencePageItem = referencePageItems[i];
        var referenceFillStroke = getFillStrokeState(referencePageItem);
        /* 片方だけ「なし」の基準は「共通 > アピアランス」で色まで照合されるので、ここでは扱わない
           A reference with only one side set to none is matched by color in Same > Appearance, so it is not handled here */
        if (referenceFillStroke && !referenceFillStroke.filled && !referenceFillStroke.stroked) {
            hasNoFillNoStrokeReference = true;
        }
        /* 前の結果に含まれていた基準は、検索しても同じ結果になる / A reference already found would give the same result */
        if (isCoveredReference[i]) continue;
        /* 基準が1つだけなら今の選択のまま実行する（選択の付け替えと redraw を省く）
           With a single reference, run on the current selection (no reselection or redraw) */
        foundGroups.push(findPageItemsWithSameAppearance(referencePageItem, referencePageItems.length === 1));
        /* 今の選択に入った後続の基準は検索を省く / Skip later references that are now selected */
        for (var j = i + 1; j < referencePageItems.length; j++) {
            if (isCoveredReference[j]) continue;
            if (usesSameFindCommand(referencePageItem, referencePageItems[j]) && referencePageItems[j].selected) {
                isCoveredReference[j] = true;
            }
        }
    }

    /* 最後の検索結果は選択されたままなので、それより前の結果だけ選択に加える
       The last result is still selected, so add only the earlier results */
    var hasAddedSelection = false;
    for (var g = 0; g < foundGroups.length - 1; g++) {
        for (var k = 0; k < foundGroups[g].length; k++) {
            foundGroups[g][k].selected = true;
            hasAddedSelection = true;
        }
    }

    /* 〈塗り〉も〈線〉もない基準があれば、同じく塗りも線もないパスをすべて加える / Add every path with neither fill nor stroke */
    if (hasNoFillNoStrokeReference) {
        var noFillStrokePageItems = findNoFillNoStrokePageItems();
        for (var m = 0; m < noFillStrokePageItems.length; m++) {
            noFillStrokePageItems[m].selected = true;
            hasAddedSelection = true;
        }
    }
    /* メニューコマンドの結果のままなら描き直さない / No redraw when the selection is the menu command's own result */
    if (hasAddedSelection) app.redraw();
})();
