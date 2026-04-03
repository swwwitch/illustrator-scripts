#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.1.7";

/*

TextScopeEditor.jsx

概要
選択した条件に応じてドキュメント内のテキストを収集し、一覧表示して内容を編集できます。
対象テキスト（アートボード）では、現在のアートボード内 / すべてのアートボード内 / ドキュメント全体を切り替えできます。
対象テキスト（レイヤー）では、//ではじめるレイヤーを含めるかどうか、ロックされたテキスト、非表示のテキストを対象に含めるかどうかを切り替えできます。
重複を除外はデフォルトでオンです。ソート、段落単位の書式保持つき置換、プレビューに対応します。
グループ内のテキストも対象です。実質的に空のテキストオブジェクトは無視します。
ロックまたは非表示のテキストを対象に含めた場合でも、一時的に編集可能な状態にしてから内容を更新し、処理後に元の状態へ戻します。
シンボル内のテキストは編集対象外ですが、一覧表示には対応します。シンボル内テキストの読み取り結果はダイアログ表示中にキャッシュされ、通常の一覧更新では再読込しません。
シンボルの再シンボル化には互換性のあるダイナミックアクション定義を使います。アクションは必要時に一度だけ読み込み、終了時に解放します。シンボル内テキスト収集では breakLink 後に毎回再シンボル化を実行しますが、再シンボル化前に選択可能なオブジェクトだけを再帰収集して選択し、シンボル化しにくい項目は除外します。不発時はアクションを再読み込みして一度だけ再試行します。
シンボル内テキスト一覧はダイアログ表示中にキャッシュしますが、初回収集時は全シンボルで再シンボル化を実行するロジックです。
OKで現在選択中のテキストに編集内容を反映して閉じます。キャンセルでダイアログを閉じます。

Overview
Collect text in the document based on the selected conditions, show it in a list, and edit the contents.
In Text Scope (Artboards), you can switch between the current artboard, all artboards, and the entire document.
In Text Scope (Layers), you can choose whether to include layers starting with //, and whether locked text and hidden text are included.
Remove Duplicates is enabled by default. Sorting, replacement while keeping paragraph-level formatting, and preview are supported.
Text inside groups is supported. Effectively empty text objects are ignored.
When locked or hidden text is included, the script temporarily makes it editable, updates the contents, and then restores the original state.
Text inside symbols is not editable, but listing is supported. Symbol text results are cached while the dialog is open, and normal list updates do not reload them.
A compatible dynamic action definition is used for re-symbolizing. The action is loaded only when needed and unloaded at the end. During symbol text collection, the script always re-symbolizes after breakLink, but before that it recursively selects only eligible objects and skips items that are unsuitable for symbol creation. If it does not fire, it reloads the action and retries once.
The symbol text list is cached while the dialog is open, but the initial collection now runs re-symbolizing for every symbol.
Pressing OK applies the current edit to the selected text and closes the dialog. Cancel closes the dialog.

更新日 / Updated: 2026-04-02

紹介記事（note）
https://note.com/dtp_tranist/n/nb845889dd553

Special thanks to:
Sergey Osokin
https://github.com/creold/illustrator-scripts/blob/master/md/Text.md#multiedittext

Released under the MIT license
http://opensource.org/licenses/mit-license.php
*/

// =========================================
// バージョンとローカライズ
// =========================================

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "テキストの収集と編集", en: "Collect and Edit Text" },
    panelTargetText: { ja: "対象テキスト（アートボード）", en: "Text Scope (Artboards)" },
    panelLayerText: { ja: "対象テキスト（レイヤー）", en: "Text Scope (Layers)" },
    rbCurrentArtboard: { ja: "現在のアートボード内", en: "Current Artboard Only" },
    rbAllArtboards: { ja: "すべてのアートボード内", en: "All Artboards Only" },
    cbDedup: { ja: "重複を除外", en: "Remove Duplicates" },
    cbOutside: { ja: "ドキュメント全体を対象に", en: "Include Entire Document" },
    cbSkipComment: { ja: "//ではじめるレイヤーも含む", en: "Include Layers That Start with //" },
    cbIncludeLocked: { ja: "ロックされたテキスト", en: "Locked Text" },
    cbIncludeHidden: { ja: "非表示のテキスト", en: "Hidden Text" },
    panelSort: { ja: "ソート", en: "Sort" },
    sortNone: { ja: "なし", en: "None" },
    sortXY: { ja: "位置順", en: "Sort by Position" },
    sortABC: { ja: "ABC順", en: "Sort Alphabetically" },
    keepFormat: { ja: "段落書式を保持", en: "Keep Paragraph Formatting" },
    keepFormatTip: { ja: "段落の書式を保持したまま\nテキストを置換します", en: "Replace text while preserving\nparagraph formatting" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    preview: { ja: "プレビュー", en: "Preview" },
    previewTip: { ja: "編集結果をリアルタイムで\nプレビューします", en: "Preview the edited result\nin real time" },
    previewDisabledTip: { ja: "CC 2020ではプレビュー無効", en: "Preview is disabled in CC 2020" },
    noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
    itemPrefix: { ja: ": ", en: ": " },
    symbolTextLabel: { ja: "シンボル内テキスト（編集不可）", en: "Text in Symbols (Read-Only)" }
};

function L(key) {
    return LABELS[key][lang];
}

function main() {
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }
    var doc = app.activeDocument;
    var isUndo = false;
    try {

        var SOFT_BREAK = '@#'; /* ソフト改行の代替文字 / Placeholder for soft line breaks */

        /* テキストフレームの参照を保持する配列 / Store references to text frames */
        var textFrameList = [];
        /* 重複除外時: duplicateMap[i] = 同じ内容を持つ全フレームの配列 / For duplicate removal: duplicateMap[i] stores all frames with the same content */
        var duplicateMap = [];

        var symbolTextCache = null;
        var symbolActionLoaded = false;
        var symbolActionFile = null;
        var SYMBOL_ACTION_SET_NAME = 'Symbol';
        var SYMBOL_ACTION_NAME = 'New';
        var TEMP_LAYER_PREFIX = '__TextScopeEdit_temp_read__';

        /* listbox用のラベルを生成（先頭部分を表示） / Build labels for the listbox using the leading part of the text */
        function makeLabel(text, maxLen) {
            if (!maxLen) maxLen = 40;
            var s = text.replace(/[\r\n]+/g, " ");
            if (s.length > maxLen) s = s.substring(0, maxLen) + "…";
            return s;
        }
        /* 実質的に空のテキストか判定 / Check whether a text frame is effectively empty */
        function isEmptyTextFrame(tf) {
            try {
                if (!tf || tf.typename !== "TextFrame") return true;
                var s = tf.contents;
                if (!s) return true;
                s = s.replace(/[\r\n\x03]/g, "").replace(/\s+/g, "");
                return s.length === 0;
            } catch (e) {
                return true;
            }
        }
        /* レイヤー名が // ではじまるか判定 / Check whether the layer name starts with // */
        function isCommentLayer(item) {
            try {
                return item.layer.name.indexOf("//") === 0;
            } catch (err) {
                return false;
            }
        }

        /* 非表示・ロック状態の判定 / Check visibility and lock state */
        function isCollectable(item, options) {
            try {
                if (!item) return false;
                if (!options.includeHidden && item.hidden) return false;
                if (!options.includeLocked && item.locked) return false;
                if (item.layer) {
                    if (!options.includeLocked && item.layer.locked) return false;
                    if (!options.includeHidden && item.layer.visible === false) return false;
                }
                return true;
            } catch (e) {
                return false;
            }
        }

        /* 全レイヤーを起点に再帰的に探索する関数 / Recursively collect items starting from top-level layers */
        function collectTextFromContainer(items, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === "TextFrame") {
                    if (isEmptyTextFrame(item)) continue;
                    textFrameList.push(item);
                } else if (item.typename === "GroupItem") {
                    collectTextFromContainer(item.pageItems, options);
                }
            }
        }

        /* アートボードと重なっているか判定する関数（一部でも重なればtrue） / Check whether an item overlaps an artboard (true even for partial overlap) */
        function isOnArtboard(item, abRect) {
            var gb = item.geometricBounds;
            return (gb[2] > abRect[0] && gb[0] < abRect[2] &&
                gb[1] > abRect[3] && gb[3] < abRect[1]);
        }

        /* アートボード内のアイテムだけ収集する関数 / Collect only items inside the target artboard */
        function collectTextOnArtboardFromContainer(items, abRect, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === "TextFrame") {
                    if (isEmptyTextFrame(item)) continue;
                    if (isOnArtboard(item, abRect)) {
                        textFrameList.push(item);
                    }
                } else if (item.typename === "GroupItem") {
                    collectTextOnArtboardFromContainer(item.pageItems, abRect, options);
                }
            }
        }

        /* いずれかのアートボードに重なっているか判定する関数 / Check whether an item overlaps any artboard */
        function isOnAnyArtboard(item) {
            for (var a = 0; a < doc.artboards.length; a++) {
                if (isOnArtboard(item, doc.artboards[a].artboardRect)) {
                    return true;
                }
            }
            return false;
        }

        /* すべてのアートボードに属するテキストを収集する関数 / Collect text that overlaps any artboard */
        function collectTextOnAllArtboardsFromContainer(items, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === "TextFrame") {
                    if (isEmptyTextFrame(item)) continue;
                    if (isOnAnyArtboard(item)) {
                        textFrameList.push(item);
                    }
                } else if (item.typename === "GroupItem") {
                    collectTextOnAllArtboardsFromContainer(item.pageItems, options);
                }
            }
        }

        function collectTextFromDocument(options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextFromContainer(doc.layers[i].pageItems, options);
            }
        }

        function collectTextOnCurrentArtboardFromDocument(abRect, options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextOnArtboardFromContainer(doc.layers[i].pageItems, abRect, options);
            }
        }

        function collectTextOnAllArtboardsFromDocument(options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextOnAllArtboardsFromContainer(doc.layers[i].pageItems, options);
            }
        }

        /* テキスト収集実行 / Run text collection */
        function gatherFrames(mode) {
            var options = {
                includeComment: cbSkipComment.value,
                includeLocked: cbIncludeLocked.value,
                includeHidden: cbIncludeHidden.value
            };
            textFrameList = [];
            duplicateMap = [];
            if (mode === "all") {
                collectTextFromDocument(options);
            } else if (mode === "allArtboards") {
                collectTextOnAllArtboardsFromDocument(options);
            } else {
                var abIndex = doc.artboards.getActiveArtboardIndex();
                var abRect = doc.artboards[abIndex].artboardRect;
                collectTextOnCurrentArtboardFromDocument(abRect, options);
            }
        }

        /* 重複を除外（同じ内容のフレームをグループ化して保持） / Remove duplicates by grouping frames with identical contents */
        function removeDuplicateFrames() {
            var unique = [];
            var groups = {}; /* contents -> index in unique array / Map contents to the index in the unique array */
            duplicateMap = [];
            for (var i = 0; i < textFrameList.length; i++) {
                var c = textFrameList[i].contents;
                if (groups[c] === undefined) {
                    groups[c] = unique.length;
                    unique.push(textFrameList[i]);
                    duplicateMap.push([textFrameList[i]]);
                } else {
                    duplicateMap[groups[c]].push(textFrameList[i]);
                }
            }
            textFrameList = unique;
        }

        // =========================================
        // 段落書式保持のためのユーティリティ / Utilities for keeping paragraph formatting
        // =========================================

        function trim(str) {
            return str.replace(/^\s+|\s+$/g, '');
        }

        function getProps(obj, keys) {
            var props = {};
            for (var i = 0; i < keys.length; i++) {
                var key = keys[i];
                try {
                    props[key] = obj[key];
                } catch (err) { }
            }
            return props;
        }

        function pasteProps(props, obj, keys) {
            for (var i = 0; i < keys.length; i++) {
                var key = keys[i];
                if (props[key] === undefined) continue;
                try {
                    obj[key] = props[key];
                } catch (err) { }
            }
        }

        function getParagraphs(obj) {
            if (!obj || !obj.story) return [];
            return obj.paragraphs;
        }

        /* 一時的に編集可能状態にして処理し、終了後に元へ戻す / Temporarily make an item editable, then restore its original state */
        function withTemporarilyEditableItem(item, fn) {
            var targetLayer = null;
            var originalItemLocked = null;
            var originalItemHidden = null;
            var originalLayerLocked = null;
            var originalLayerVisible = null;

            try {
                if (!item) return;
                targetLayer = item.layer ? item.layer : null;

                originalItemLocked = item.locked;
                originalItemHidden = item.hidden;
                if (targetLayer) {
                    originalLayerLocked = targetLayer.locked;
                    originalLayerVisible = targetLayer.visible;
                }

                if (targetLayer) {
                    if (targetLayer.visible === false) targetLayer.visible = true;
                    if (targetLayer.locked) targetLayer.locked = false;
                }
                if (item.hidden) item.hidden = false;
                if (item.locked) item.locked = false;

                fn();
            } finally {
                try {
                    if (item && originalItemLocked !== null) item.locked = originalItemLocked;
                } catch (e) { }
                try {
                    if (item && originalItemHidden !== null) item.hidden = originalItemHidden;
                } catch (e) { }
                try {
                    if (targetLayer && originalLayerLocked !== null) targetLayer.locked = originalLayerLocked;
                } catch (e) { }
                try {
                    if (targetLayer && originalLayerVisible !== null) targetLayer.visible = originalLayerVisible;
                } catch (e) { }
            }
        }

        /* ローカル書式は保持せず、段落単位の書式のみ保持 / Keep paragraph-level formatting only; do not preserve local character formatting */
        function replaceContent(tf, str, isKeepStyle) {
            if (!/text/i.test(tf.typename)) return;

            withTemporarilyEditableItem(tf, function () {
                if (!isKeepStyle) {
                    tf.contents = str;
                    return;
                }

                var paras = getParagraphs(tf);
                var styles = [];
                var paraKeys = [
                    'justification', 'firstLineIndent', 'leftIndent', 'rightIndent',
                    'spaceBefore', 'spaceAfter', 'hyphenation', 'hyphenationZone',
                    'desiredWordSpacing', 'minimumWordSpacing', 'maximumWordSpacing',
                    'desiredLetterSpacing', 'minimumLetterSpacing', 'maximumLetterSpacing',
                    'desiredGlyphScaling', 'minimumGlyphScaling', 'maximumGlyphScaling',
                    'singleWordJustification', 'everyLineComposer', 'kinsokuOrder',
                    'bunriKinshi', 'kurikaeshiMojiShori', 'romanHanging', 'mojikumi',
                    'kinsoku', 'leadingType'
                ];

                try {
                    for (var i = 0; i < paras.length; i++) {
                        styles.push(getProps(paras[i], paraKeys));
                    }
                } catch (err) { }

                tf.contents = str;
                paras = getParagraphs(tf);

                if (styles.length === 0) {
                    return;
                }

                var style = null;
                for (var j = 0; j < paras.length; j++) {
                    style = styles[j] ? styles[j] : styles[styles.length - 1];
                    pasteProps(style, paras[j], paraKeys);
                }
            });
        }

        // =========================================
        // ソート / Sorting
        // =========================================

        function sortByPosition(coll, tolerance) {
            if (!tolerance) tolerance = 10;
            coll.sort(function (a, b) {
                if (Math.abs(b.top - a.top) <= tolerance) {
                    return a.left - b.left;
                }
                return b.top - a.top;
            });
        }

        function sortByContent(coll) {
            coll.sort(function (a, b) {
                var ca = a.contents.toLowerCase();
                var cb = b.contents.toLowerCase();
                if (ca < cb) return -1;
                if (ca > cb) return 1;
                return 0;
            });
        }

        // =========================================
        function ensureNewSymbolActionLoaded(forceReload) {
            if (forceReload) {
                unloadNewSymbolAction();
            }
            if (symbolActionLoaded) return;

            var str = '/version 3\n' +
                '/name [ 6\n' +
                '\t53796d626f6c\n' +
                ']\n' +
                '/isOpen 1\n' +
                '/actionCount 1\n' +
                '/action-1 {\n' +
                '\t/name [ 3\n' +
                '\t\t4e6577\n' +
                '\t]\n' +
                '\t/keyIndex 0\n' +
                '\t/colorIndex 0\n' +
                '\t/isOpen 1\n' +
                '\t/eventCount 1\n' +
                '\t/event-1 {\n' +
                '\t\t/useRulersIn1stQuadrant 0\n' +
                '\t\t/internalName (ai_plugin_symbol_palette)\n' +
                '\t\t/localizedName [ 12\n' +
                '\t\t\te382b7e383b3e3839ce383ab\n' +
                '\t\t]\n' +
                '\t\t/isOpen 1\n' +
                '\t\t/isOn 1\n' +
                '\t\t/hasDialog 1\n' +
                '\t\t/showDialog 0\n' +
                '\t\t/parameterCount 2\n' +
                '\t\t/parameter-1 {\n' +
                '\t\t\t/key 1835363957\n' +
                '\t\t\t/showInPalette 4294967295\n' +
                '\t\t\t/type (enumerated)\n' +
                '\t\t\t/name [ 18\n' +
                '\t\t\t\te696b0e8a68fe382b7e383b3e3839ce383ab\n' +
                '\t\t\t]\n' +
                '\t\t\t/value 2\n' +
                '\t\t}\n' +
                '\t\t/parameter-2 {\n' +
                '\t\t\t/key 1919250540\n' +
                '\t\t\t/showInPalette 4294967295\n' +
                '\t\t\t/type (boolean)\n' +
                '\t\t\t/value 1\n' +
                '\t\t}\n' +
                '\t}\n' +
                '}\n';

            var f = new File(Folder.temp + '/TextScopeEdit_SymbolAction.aia');
            f.open('w');
            f.write(str);
            f.close();
            app.loadAction(f);
            symbolActionLoaded = true;
            symbolActionFile = f;
        }

        function unloadNewSymbolAction() {
            if (!symbolActionLoaded) return;
            try {
                app.unloadAction(SYMBOL_ACTION_SET_NAME, '');
            } catch (e) { }
            try {
                if (symbolActionFile && symbolActionFile.exists) symbolActionFile.remove();
            } catch (e2) { }
            symbolActionLoaded = false;
            symbolActionFile = null;
        }

        function getSelectionArray() {
            var arr = [];
            try {
                for (var i = 0; i < doc.selection.length; i++) {
                    arr.push(doc.selection[i]);
                }
            } catch (e) { }
            return arr;
        }

        function selectionHasSymbolItem(selectionArray) {
            for (var i = 0; i < selectionArray.length; i++) {
                if (selectionArray[i] && selectionArray[i].typename === 'SymbolItem') {
                    return true;
                }
            }
            return false;
        }

        function isResymbolizableItem(item) {
            if (!item) return false;
            try {
                if (item.locked || item.hidden) return false;
            } catch (e) { }
            if (item.guides) return false;
            if (item.clipping) return false;
            if (item.typename === 'PluginItem' || item.typename === 'NonNativeItem' || item.typename === 'LegacyTextItem') {
                return false;
            }
            return !!item.selected;
        }

        function collectResymbolizableItems(container, out) {
            if (!container || !container.pageItems) return;
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === 'GroupItem' || item.typename === 'CompoundPathItem') {
                    if (isResymbolizableItem(item)) {
                        out.push(item);
                    } else {
                        collectResymbolizableItems(item, out);
                    }
                } else if (isResymbolizableItem(item)) {
                    out.push(item);
                }
            }
        }

        function selectResymbolizableItemsInContainer(container) {
            var items = [];
            collectResymbolizableItems(container, items);
            try {
                doc.selection = null;
            } catch (e) { }
            for (var i = 0; i < items.length; i++) {
                try {
                    items[i].selected = true;
                } catch (e2) { }
            }
            return items.length;
        }

        function act_NewSymbol() {
            ensureNewSymbolActionLoaded(false);
            app.doScript(SYMBOL_ACTION_NAME, SYMBOL_ACTION_SET_NAME, false);
            var selectionAfter = getSelectionArray();
            if (!selectionHasSymbolItem(selectionAfter)) {
                ensureNewSymbolActionLoaded(true);
                app.doScript(SYMBOL_ACTION_NAME, SYMBOL_ACTION_SET_NAME, false);
            }
        }

        function createTempReadLayer() {
            var layer = doc.layers.add();
            layer.name = TEMP_LAYER_PREFIX + (new Date().getTime());
            return layer;
        }
        // シンボル内テキスト収集 / Collect text inside symbols
        // =========================================

        function collectSymbolItemsDeep(container, out) {
            if (!container || !container.pageItems) return;
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "SymbolItem") {
                    out.push(item);
                }
                if (item.pageItems) {
                    collectSymbolItemsDeep(item, out);
                }
            }
        }

        function expandAllSymbolsInContainer(container) {
            if (!container) return;
            while (true) {
                var symbols = [];
                collectSymbolItemsDeep(container, symbols);
                if (symbols.length === 0) break;
                for (var i = symbols.length - 1; i >= 0; i--) {
                    try {
                        symbols[i].breakLink();
                    } catch (e) { }
                }
            }
        }

        function searchSymbolText(container, symName, results) {
            if (!container) return;
            if (container.typename === "TextFrame") {
                if (!isEmptyTextFrame(container)) {
                    results.push(symName + "：" + container.contents.replace(/[\r\n]+/g, " "));
                }
                return;
            }
            if (!container.pageItems) return;
            for (var j = 0; j < container.pageItems.length; j++) {
                var item = container.pageItems[j];
                if (item.typename === "TextFrame") {
                    if (!isEmptyTextFrame(item)) {
                        results.push(symName + "：" + item.contents.replace(/[\r\n]+/g, " "));
                    }
                } else if (item.pageItems) {
                    searchSymbolText(item, symName, results);
                }
            }
        }

        function collectSymbolTextsFromTempLayer(tempLayer, symName) {
            var results = [];
            searchSymbolText(tempLayer, symName, results);
            return results;
        }

        function withSelectionRestored(fn) {
            var prevSelection = [];
            try {
                for (var i = 0; i < doc.selection.length; i++) {
                    prevSelection.push(doc.selection[i]);
                }
            } catch (e) { }

            try {
                fn();
            } finally {
                try {
                    doc.selection = null;
                } catch (e2) { }
                for (var j = 0; j < prevSelection.length; j++) {
                    try {
                        prevSelection[j].selected = true;
                    } catch (e3) { }
                }
            }
        }

        function collectSymbolTexts() {
            var results = [];
            var symbolItems = doc.symbolItems;
            var siList = [];
            for (var i = 0; i < symbolItems.length; i++) {
                siList.push(symbolItems[i]);
            }
            var processed = {};

            withSelectionRestored(function () {
                for (var i = 0; i < siList.length; i++) {
                    var symName = siList[i].symbol.name;
                    if (processed[symName]) continue;
                    processed[symName] = true;

                    var tempLayer = createTempReadLayer();
                    try {
                        var placedItem = tempLayer.symbolItems.add(siList[i].symbol);
                        if (placedItem && placedItem.typename === "SymbolItem") {
                            try {
                                placedItem.breakLink();
                            } catch (e) { }
                        }
                        expandAllSymbolsInContainer(tempLayer);
                        var firstPass = collectSymbolTextsFromTempLayer(tempLayer, symName);
                        for (var r1 = 0; r1 < firstPass.length; r1++) {
                            results.push(firstPass[r1]);
                        }
                        if (tempLayer.pageItems.length > 0) {
                            var selectedCount = selectResymbolizableItemsInContainer(tempLayer);
                            if (selectedCount > 0) {
                                act_NewSymbol();

                                var sel = getSelectionArray();
                                for (var m = sel.length - 1; m >= 0; m--) {
                                    if (sel[m] && sel[m].typename === "SymbolItem") {
                                        try {
                                            sel[m].breakLink();
                                        } catch (e4) { }
                                    }
                                }
                                expandAllSymbolsInContainer(tempLayer);

                                var secondPass = collectSymbolTextsFromTempLayer(tempLayer, symName);
                                for (var r2 = 0; r2 < secondPass.length; r2++) {
                                    results.push(secondPass[r2]);
                                }
                            }
                        }
                    } catch (e5) {
                    } finally {
                        try {
                            doc.selection = null;
                        } catch (e6) { }
                        try {
                            tempLayer.remove();
                        } catch (e7) { }
                    }
                }
            });

            return results;
        }

        // =========================================
        // listboxを更新 / Update listbox
        // =========================================

        function updateList() {
            var mode;
            if (cbOutside.value) {
                mode = "all";
            } else if (rbAll.value) {
                mode = "allArtboards";
            } else {
                mode = "current";
            }
            gatherFrames(mode);

            /* ソート / Sorting */
            if (rbSortXY.value) {
                sortByPosition(textFrameList);
            } else if (rbSortABC.value) {
                sortByContent(textFrameList);
            }

            if (cbDedup.value) {
                removeDuplicateFrames();
            }

            listBox.removeAll();
            for (var i = 0; i < textFrameList.length; i++) {
                listBox.add("item", (i + 1) + L("itemPrefix") + makeLabel(textFrameList[i].contents));
            }
            editBox.text = "";
            if (textFrameList.length > 0) {
                listBox.selection = 0;
            }
        }

        function getSymbolTexts(forceRefresh) {
            if (!forceRefresh && symbolTextCache !== null) {
                return symbolTextCache;
            }
            symbolTextCache = collectSymbolTexts();
            return symbolTextCache;
        }

        function refreshSymbolList(forceRefresh) {
            symbolListBox.removeAll();
            var symTexts = getSymbolTexts(forceRefresh);
            for (var s = 0; s < symTexts.length; s++) {
                symbolListBox.add("item", symTexts[s]);
            }
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        function buildDialogUI() {
            var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];

            /* 2カラム: 左=listbox+editbox、右=オプション、下=ボタンエリア / Two columns: left = listbox + edit box, right = options, bottom = button area */
            var mainGroup = dlg.add("group");
            mainGroup.orientation = "row";
            mainGroup.alignChildren = ["fill", "fill"];

            /* 左カラム / Left column */
            var leftCol = mainGroup.add("group");
            leftCol.orientation = "column";
            leftCol.alignChildren = ["fill", "top"];


            var listBox = leftCol.add("listbox", [0, 0, 250, 230], []);
            var editBox = leftCol.add("edittext", [0, 0, 250, 50], "", { multiline: true, scrolling: true });

            /* シンボル内テキスト / Text in symbols */
            leftCol.add("statictext", undefined, L("symbolTextLabel"));
            var symbolListBox = leftCol.add("listbox", [0, 0, 250, 50], []);
            symbolListBox.alignment = ["fill", "fill"];

            /* 右カラム / Right column */
            var rightCol = mainGroup.add("group");
            rightCol.orientation = "column";
            rightCol.alignChildren = ["fill", "top"];

            /* 対象テキストパネル / Target text panel */
            var targetPanel = rightCol.add("panel", undefined, L("panelTargetText"));
            targetPanel.orientation = "column";
            targetPanel.alignChildren = ["left", "top"];
            targetPanel.margins = [15, 20, 15, 10];

            var rbGroup = targetPanel.add("group");
            rbGroup.orientation = "column";
            rbGroup.alignChildren = ["left", "top"];
            var rbArtboard = rbGroup.add("radiobutton", undefined, L("rbCurrentArtboard"));
            var rbAll = rbGroup.add("radiobutton", undefined, L("rbAllArtboards"));
            rbArtboard.value = true;

            var cbRow = targetPanel.add("group");
            cbRow.orientation = "row";
            var cbOutside = cbRow.add("checkbox", undefined, L("cbOutside"));
            cbOutside.value = false;
            cbOutside.enabled = false;

            var layerPanel = rightCol.add("panel", undefined, L("panelLayerText"));
            layerPanel.orientation = "column";
            layerPanel.alignChildren = ["left", "top"];
            layerPanel.margins = [15, 20, 15, 10];

            var cbSkipComment = layerPanel.add("checkbox", undefined, L("cbSkipComment"));
            cbSkipComment.value = false;

            var cbIncludeLocked = layerPanel.add("checkbox", undefined, L("cbIncludeLocked"));
            cbIncludeLocked.value = false;
            var cbIncludeHidden = layerPanel.add("checkbox", undefined, L("cbIncludeHidden"));
            cbIncludeHidden.value = false;

            var sortPanel = rightCol.add("panel", undefined, L("panelSort"));
            sortPanel.orientation = "column";
            sortPanel.alignChildren = ["left", "top"];
            sortPanel.margins = [15, 20, 15, 10];
            var sortGroup = sortPanel.add("group");
            sortGroup.orientation = "row";
            sortGroup.alignChildren = ["left", "center"];
            var rbSortNone = sortGroup.add("radiobutton", undefined, L("sortNone"));
            var rbSortXY = sortGroup.add("radiobutton", undefined, L("sortXY"));
            var rbSortABC = sortGroup.add("radiobutton", undefined, L("sortABC"));
            rbSortNone.value = true;

            var cbDedup = rightCol.add("checkbox", undefined, L("cbDedup"));
            cbDedup.value = true;

            var isFormat = rightCol.add("checkbox", undefined, L("keepFormat"));
            isFormat.value = true;
            isFormat.helpTip = L("keepFormatTip");

            var buttonRow = dlg.add("group");
            buttonRow.orientation = "row";
            buttonRow.alignChildren = ["fill", "center"];

            var leftButtons = buttonRow.add("group");
            leftButtons.orientation = "row";
            leftButtons.alignChildren = ["left", "center"];
            var isPreview = leftButtons.add("checkbox", undefined, L("preview"));
            isPreview.helpTip = L("previewTip");

            var spacer = buttonRow.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var rightButtons = buttonRow.add("group");
            rightButtons.orientation = "row";
            rightButtons.alignChildren = ["right", "center"];
            var cancelBtn = rightButtons.add("button", undefined, L("cancel"), { name: "cancel" });
            var closeBtn = rightButtons.add("button", undefined, L("ok"), { name: "ok" });
            dlg.defaultElement = closeBtn;

            /* CC 2020 v24.3 はプレビュー時にクラッシュするため無効化 / Disable preview in CC 2020 v24.3 because it may crash */
            if (parseInt(app.version) == 24) {
                isPreview.enabled = false;
                isPreview.helpTip = L("previewDisabledTip");
            }

            return {
                dlg: dlg,
                rbArtboard: rbArtboard,
                rbAll: rbAll,
                cbDedup: cbDedup,
                cbOutside: cbOutside,
                cbSkipComment: cbSkipComment,
                cbIncludeLocked: cbIncludeLocked,
                cbIncludeHidden: cbIncludeHidden,
                listBox: listBox,
                editBox: editBox,
                rbSortNone: rbSortNone,
                rbSortXY: rbSortXY,
                rbSortABC: rbSortABC,
                isFormat: isFormat,
                cancelBtn: cancelBtn,
                closeBtn: closeBtn,
                isPreview: isPreview,
                symbolListBox: symbolListBox
            };
        }

        var ui = buildDialogUI();
        var dlg = ui.dlg;
        var rbArtboard = ui.rbArtboard;
        var rbAll = ui.rbAll;
        var cbDedup = ui.cbDedup;
        var cbOutside = ui.cbOutside;
        var cbSkipComment = ui.cbSkipComment;
        var cbIncludeLocked = ui.cbIncludeLocked;
        var cbIncludeHidden = ui.cbIncludeHidden;
        var listBox = ui.listBox;
        var editBox = ui.editBox;
        var rbSortNone = ui.rbSortNone;
        var rbSortXY = ui.rbSortXY;
        var rbSortABC = ui.rbSortABC;
        var isFormat = ui.isFormat;
        var cancelBtn = ui.cancelBtn;
        var closeBtn = ui.closeBtn;
        var isPreview = ui.isPreview;
        var symbolListBox = ui.symbolListBox;

        // =========================================
        // イベントハンドラ / Event handlers
        // =========================================

        /* listbox選択時にeditboxを更新 / Update the edit box when the listbox selection changes */
        listBox.onChange = function () {
            if (listBox.selection !== null) {
                var idx = listBox.selection.index;
                editBox.text = textFrameList[idx].contents.replace(/\x03/g, SOFT_BREAK);
            }
        };

        function reflectPreviewEnabledState() {
            if (parseInt(app.version) != 24) {
                isPreview.enabled = !isFormat.value;
            }
        }

        /* 段落書式保持とプレビューの排他制御 / Make paragraph formatting and preview mutually exclusive */
        isFormat.onClick = function () {
            reflectPreviewEnabledState();
        };

        /* Shift+Enter でソフト改行文字を挿入 / Insert the soft-break placeholder with Shift+Enter */
        editBox.addEventListener('keydown', function (kd) {
            var isShift = ScriptUI.environment.keyboardState['shiftKey'];
            if (isShift && kd.keyName === 'Enter') {
                this.textselection = SOFT_BREAK;
                kd.preventDefault();
            }
        });

        /* プレビュー / Preview */
        function preview() {
            if (parseInt(app.version) == 24) return;
            try {
                if (isPreview.enabled && isPreview.value && listBox.selection !== null) {
                    if (isUndo) app.undo();
                    else isUndo = true;
                    applyCurrentEdit();
                    app.redraw();
                } else if (isUndo) {
                    app.undo();
                    app.redraw();
                    isUndo = false;
                }
            } catch (err) { }
        }

        editBox.onChanging = function () { preview(); };
        isPreview.onClick = function () { preview(); };

        /* 選択中のテキストフレームに編集を反映（重複があれば全フレームに適用） / Apply the current edit to the selected text frame, or all duplicates when enabled */
        function applyCurrentEdit() {
            if (listBox.selection === null) return;
            var idx = listBox.selection.index;
            var newText = editBox.text.replace(new RegExp(SOFT_BREAK, 'gmi'), '\x03');
            if (cbDedup.value && duplicateMap[idx]) {
                for (var d = 0; d < duplicateMap[idx].length; d++) {
                    replaceContent(duplicateMap[idx][d], newText, isFormat.value);
                }
            } else {
                replaceContent(textFrameList[idx], newText, isFormat.value);
            }
        }




        /* 初回収集 / Initial collection */
        updateList();
        refreshSymbolList(false);

        /* ラジオボタン・チェックボックス切り替え時に更新 / Refresh when radio buttons or checkboxes change */
        rbArtboard.onClick = function () {
            cbOutside.enabled = false;
            cbOutside.value = false;
            updateList();
        };
        rbAll.onClick = function () {
            cbOutside.enabled = true;
            updateList();
        };
        cbDedup.onClick = function () { updateList(); };


        cbOutside.onClick = function () { updateList(); };
        cbSkipComment.onClick = function () { updateList(); };
        cbIncludeLocked.onClick = function () { updateList(); };
        cbIncludeHidden.onClick = function () { updateList(); };
        rbSortNone.onClick = function () { updateList(); };
        rbSortXY.onClick = function () { updateList(); };
        rbSortABC.onClick = function () { updateList(); };

        /* キャンセルボタンで閉じる / Close the dialog when Cancel is pressed */
        cancelBtn.onClick = function () {
            dlg.close();
        };

        /* OKボタンで現在の編集を反映して閉じる / Apply the current edit and close when OK is pressed */
        closeBtn.onClick = function () {
            if (isUndo && isPreview.value) {
                app.undo();
                isUndo = false;
            }
            applyCurrentEdit();
            dlg.close();
        };

        /* ダイアログを閉じるとき、プレビュー中ならundoで元に戻す / Undo the preview when the dialog closes */
        dlg.onClose = function () {
            try {
                if (isUndo) app.undo();
                isUndo = false;
            } catch (err) { }
        };

        dlg.show();
    } finally {
        if (isUndo) {
            try {
                app.undo();
            } catch (err) { }
            isUndo = false;
        }
        unloadNewSymbolAction();
    }
}

main();