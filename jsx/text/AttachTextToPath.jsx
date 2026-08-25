#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字とパスを選択して実行すると、パス上文字に変換します。

詳細は README を参照してください。

### Overview

Converts a selected point text and path into text on a path.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AttachTextToPath";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AttachTextToPath.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AttachTextToPath.md
var SCRIPT_ARTICLE_URL = "https://note.com/gautt/n/n92f6faeda048"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // --- Localization (ja/en) ---
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var lang = getCurrentLang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noTargetText: { ja: "対象のポイント文字が見つかりません", en: "No target point text found." },
            needPath: { ja: "パスを一緒に選択してください", en: "Please select a path together with the text." },
            duplicateFailed: { ja: "パスの複製に失敗しました", en: "Failed to duplicate the path." }
        }
    };

    function L(path) {
        var parts = String(path).split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return path;
            node = node[parts[i]];
        }
        if (node == null) return path;
        if (node[lang] != null) return node[lang];
        return (node.en != null) ? node.en : path;
    }

    // Get items
    if (app.documents.length === 0) {
        alert(L('alert.noDocument'));
        return false;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;

    var targetItems = getTargetTextItems(sel);
    var selectedPaths = getSelectedPathItems(sel);

    // Validation
    if (targetItems.length === 0) {
        alert(L('alert.noTargetText'));
        return false;
    }

    if (!selectedPaths || selectedPaths.length === 0) {
        alert(L('alert.needPath'));
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
                alert(L('alert.needPath'));
                continue;
            }

            // Track used paths for removal (avoid duplicates)
            pushUnique(usedPaths, basePath);

            var textPath = duplicatePathForText(basePath, currentLayer);
            if (!textPath) {
                alert(L('alert.duplicateFailed'));
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
        } catch (_) { }
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
        } catch (_) { }
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
