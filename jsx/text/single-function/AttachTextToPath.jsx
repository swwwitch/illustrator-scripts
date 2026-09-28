#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字とパスを選択して実行すると、パス上文字に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AttachTextToPath.md

note記事も参照してください。
https://note.com/gautt/n/n92f6faeda048

### Overview

Converts a selected point text and path into text on a path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AttachTextToPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AttachTextToPath";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AttachTextToPath.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AttachTextToPath.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/gautt/n/n92f6faeda048"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // --- Localization (ja/en) ---
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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noTargetText: { ja: "対象のポイント文字が見つかりません", en: "No target point text found." },
            needPath: { ja: "パスを一緒に選択してください", en: "Please select a path together with the text." },
            duplicateFailed: { ja: "パスの複製に失敗しました", en: "Failed to duplicate the path." }
        }
    };

    // Get items
    if (app.documents.length === 0) {
        alert(getLabel('alert.noDocument'));
        return false;
    }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;

    var targetItems = getTargetTextItems(currentSelection);
    var selectedPaths = getSelectedPathItems(currentSelection);

    // Validation
    if (targetItems.length === 0) {
        alert(getLabel('alert.noTargetText'));
        return false;
    }

    if (!selectedPaths || selectedPaths.length === 0) {
        alert(getLabel('alert.needPath'));
        return false;
    }

    main(targetItems, selectedPaths);

    // Main process
    function main(targetItems, selectedPaths) {

        // Track original selected paths for later removal (avoid duplicates).
        // NOTE:
        // - Only the *original selected paths* are removed at the end.
        // - The duplicated path used for textPath remains in the document.
        var usedPaths = [];

        for (var j = 0; j < targetItems.length; j++) {
            // Get original text frame item & current layer
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // Resolve base path + duplicate path for this text
            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, j);
            if (!basePath) {
                // Should not happen because of Validation, but keep defensive
                alert(getLabel('alert.needPath'));
                continue;
            }

            // Track used paths for removal (avoid duplicates)
            pushUnique(usedPaths, basePath);

            var textPath = duplicatePathForText(basePath, currentLayer);
            if (!textPath) {
                alert(getLabel('alert.duplicateFailed'));
                continue;
            }

            // Create Text on a path
            var textOnAPath = currentLayer.textFrames.pathText(textPath);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // Remove original text frame item
            safeRemove(originalText);

            // Select text on a path
            textOnAPath.selected = true;
        }

        // Remove only the original selected paths (not the duplicated textPath).
        // The duplicated paths created for path text must remain.
        for (var p = 0; p < usedPaths.length; p++) {
            safeRemove(usedPaths[p]);
        }
    }

    function safeRemove(item) {
        try {
            if (!item) return;
            item.remove();
        } catch (e) { }
    }

    // Push only if not already in the array (by reference)
    function pushUnique(arr, item) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === item) return;
        }
        arr.push(item);
    }

    // Decide which selected path to use for each text.
    // Rule 1: If the number of selected paths equals the number of texts,
    //         use the path at the same index (1-to-1 correspondence).
    // Rule 2: Otherwise, use the first selected path for all texts.
    function resolveBasePathForText(selectedPaths, textCount, index) {
        if (!selectedPaths || selectedPaths.length === 0) return null;
        if (selectedPaths.length === textCount) return selectedPaths[index];
        return selectedPaths[0];
    }

    // Resolve actual PathItem to duplicate (CompoundPathItem -> first pathItem)
    function resolveSourcePath(basePath) {
        if (!basePath) return null;

        if (basePath.typename === 'CompoundPathItem') {
            if (basePath.pathItems && basePath.pathItems.length > 0) {
                return basePath.pathItems[0];
            }
            return null;
        }

        return basePath;
    }

    // Duplicate a path for the text and place it on the same layer
    function duplicatePathForText(basePath, currentLayer) {
        var srcPath = resolveSourcePath(basePath);
        if (!srcPath) return null;

        try {
            var dup = srcPath.duplicate();
            tryMoveToLayer(dup, currentLayer);
            return dup;
        } catch (eDup) {
            return null;
        }
    }

    function tryMoveToLayer(item, layer) {
        try {
            if (!item || !layer) return;
            item.move(layer, ElementPlacement.PLACEATBEGINNING);
        } catch (e) { }
    }

    // Collect items recursively from GroupItem.pageItems using a predicate.
    // Why only GroupItem recursion?
    // - In Illustrator's object model, GroupItem is the general container that nests pageItems.
    // - PathItem and CompoundPathItem are collected as-is (CompoundPathItem is handled later when duplicating).
    // - Keeping recursion limited to GroupItem avoids unintended traversal into structures that behave differently
    //   (e.g. clipping/compound internals) and keeps selection-based behavior predictable.
    //
    // - items: collection/array of pageItems
    // - acceptFn: function(item) -> boolean
    function collectRecursive(items, acceptFn) {
        var out = [];
        if (!items) return out;

        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;

            if (acceptFn && acceptFn(it)) {
                out.push(it);
            } else if (it.typename === 'GroupItem') {
                out = out.concat(collectRecursive(it.pageItems, acceptFn));
            }
        }
        return out;
    }

    // Get target point-text items (recursive)
    function getTargetTextItems(items) {
        return collectRecursive(items, function (it) {
            return (it.typename === 'TextFrame' && it.kind === TextType.POINTTEXT);
        });
    }

    // Get selected path items (PathItem or CompoundPathItem). If GroupItem contains paths, collect them too.
    function getSelectedPathItems(items) {
        return collectRecursive(items, function (it) {
            return (it.typename === 'PathItem' || it.typename === 'CompoundPathItem');
        });
    }

}());
