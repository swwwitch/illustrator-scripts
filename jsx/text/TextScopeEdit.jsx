#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.2.6";

/*

TextScopeEdit.jsx

概要
選択した条件に応じてドキュメント内のテキストを収集し、一覧表示して内容を編集できます。
対象テキスト（アートボード）では、現在のアートボード内 / すべてのアートボード内 / ドキュメント全体を切り替えできます。
対象テキスト（レイヤー）では、//ではじめるレイヤーを含めるかどうか、ロックされたテキスト、非表示のテキストを対象に含めるかどうかを切り替えできます。
同じ内容のテキストをまとめて扱うはデフォルトでオンです。同じ内容のテキストが複数ある場合、一覧では1件として扱い、編集時は該当テキストを一括で更新します。ソート、段落単位の書式保持つき置換、プレビューに対応します。
グループ内およびサブレイヤー内のテキストも対象です。実質的に空のテキストオブジェクトは無視します。
ロックまたは非表示のテキストを対象に含めた場合でも、一時的に編集可能な状態にしてから内容を更新し、処理後に元の状態へ戻します。
シンボル内のテキストは編集対象外ですが、一覧表示には対応します。シンボル内テキストの読み取り結果はダイアログ表示中にキャッシュされ、通常の一覧更新では再読込しません。
シンボル内テキストの読み取りは、同名シンボルを1回だけ対象にします。対象シンボルの代表インスタンスを一時レイヤーへ複製し、複製側だけを breakLink して展開後のテキストを収集します。読み取りに使った複製オブジェクトは削除し、元のカンバス上のシンボルには触れません。一時レイヤーは必要時のみ作成し、このスクリプトが作成した場合のみ削除します。
ダイアログ左側には「テキスト一覧」「テキスト編集」「シンボル内テキスト（編集不可）」を表示します。テキスト編集欄の高さは4行分です。シンボル内テキスト一覧の高さは内容件数に応じて自動調整され、最小4行・最大8行の範囲で表示します。
段落書式を保持がオンのときは、プレビューを自動でオフにして無効化します。プレビュー中に切り替えた場合も、プレビュー状態を解除してから反映します。
［テキスト書き出し］で、ドキュメント内のすべての通常テキストとシンボル内テキストをアートボードごとにまとめてデスクトップへテキストファイルとして書き出します。書き出しは現在の一覧表示条件や重複除外、ソート状態に影響されません。各アートボードは ---アートボード番号: アートボード名--- / アートボード外は ---アートボード外--- の見出しで区切られ、シンボル内テキストも配置されているアートボードごとにまとめて出力します。シンボル内テキストは「（シンボル: 名前）」の形式で末尾に付加されます。OKで現在選択中のテキストに編集内容を反映して閉じます。キャンセルでダイアログを閉じます。

Overview
Collect text in the document based on the selected conditions, show it in a list, and edit the contents.
In Text Scope (Artboards), you can switch between the current artboard, all artboards, and the entire document.
In Text Scope (Layers), you can choose whether to include layers starting with //, and whether locked text and hidden text are included.
Treat Same Text as One is enabled by default. When multiple text frames have the same content, the list shows them as one item and editing updates all matching text frames together. Sorting, replacement while keeping paragraph-level formatting, and preview are supported.
Text inside groups and sublayers is supported. Effectively empty text objects are ignored.
When locked or hidden text is included, the script temporarily makes it editable, updates the contents, and then restores the original state.
Text inside symbols is not editable, but listing is supported. Symbol text results are cached while the dialog is open, and normal list updates do not reload them.
To read text inside symbols, the script targets each symbol name only once. It duplicates one representative instance of the target symbol to a temporary layer, break-links only the duplicate, and collects the resulting text. The temporary duplicate objects are then removed, and the original symbols on the canvas remain untouched. The temporary layer is created only when needed and is removed only if this script created it.
The left side of the dialog shows Text List, Edit Text, and Text in Symbols (Read-Only). The Edit Text field height is 4 lines. The height of the symbol text list adjusts automatically to the number of items, with a minimum of 4 rows and a maximum of 8 rows.
When Keep Paragraph Formatting is enabled, Preview is turned off automatically and disabled. If you switch it on while previewing, the preview state is cleared before the change is applied.
[Export Text] writes all regular text in the document and the text found inside symbols to a text file on the Desktop, grouped by artboard. Export is not affected by the current list view conditions, duplicate removal, or sort state. Each artboard is written under a ---Artboard Number: Artboard Name--- heading, and items outside all artboards are written under ---Outside Artboards---. Text found inside symbols is also grouped under the artboard where that symbol instance is placed. Symbol text lines append "(Symbol: name)" in English. Pressing OK applies the current edit to the selected text and closes the dialog. Cancel closes the dialog.

更新日 / Updated: 2026-04-03

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
    cbDedup: { ja: "同じ内容を一括編集", en: "Treat Same Text as One" },
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
    exportText: { ja: "テキスト書き出し", en: "Export Text" },
    exportDone: { ja: "テキストを書き出しました", en: "Text exported" },
    exportFailed: { ja: "テキストを書き出せませんでした", en: "Failed to export text" },
    preview: { ja: "プレビュー", en: "Preview" },
    previewTip: { ja: "編集結果をリアルタイムで\nプレビューします", en: "Preview the edited result\nin real time" },
    previewDisabledTip: { ja: "CC 2020ではプレビュー無効", en: "Preview is disabled in CC 2020" },
    noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
    itemPrefix: { ja: ": ", en: ": " },
    symbolTextLabel: { ja: "シンボル内テキスト（編集不可）", en: "Text in Symbols (Read-Only)" },
    textListLabel: { ja: "テキスト一覧", en: "Text List" },
    textEditLabel: { ja: "テキスト編集", en: "Edit Text" }
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
        var TEMP_LAYER_NAME = '__TextScopeEdit_temp_read__';
        var TEMP_LAYER_NOTE = '__TextScopeEdit_temp_read__';

        var SYMBOL_LIST_MIN_ROWS = 4;
        var SYMBOL_LIST_MAX_ROWS = 8;
        var SYMBOL_LIST_ROW_HEIGHT = 18;
        var SYMBOL_LIST_EXTRA_HEIGHT = 6;

        /* listbox用のラベルを生成（先頭部分を表示） / Build labels for the listbox using the leading part of the text */
        function makeLabel(text, maxLen) {
            if (!maxLen) maxLen = 40;
            var s = text.replace(/[\r\n]+/g, " ");
            if (s.length > maxLen) s = s.substring(0, maxLen) + "…";
            return s;
        }

        function zeroPad2(num) {
            return (num < 10 ? '0' : '') + num;
        }

        function getDateStamp() {
            var now = new Date();
            return now.getFullYear() + zeroPad2(now.getMonth() + 1) + zeroPad2(now.getDate());
        }

        function getDocumentBaseName(documentRef) {
            var name = documentRef && documentRef.name ? documentRef.name : 'untitled';
            return name.replace(/\.[^\.]+$/, '');
        }

        function sanitizeFileName(name) {
            return name.replace(/[\\\/\:\*\?\"\<\>\|]+/g, '_');
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

        function collectTextFromLayer(layer, options) {
            if (!layer) return;
            collectTextFromContainer(layer.pageItems, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectTextFromLayer(layer.layers[i], options);
            }
        }

        function collectTextOnCurrentArtboardFromLayer(layer, abRect, options) {
            if (!layer) return;
            collectTextOnArtboardFromContainer(layer.pageItems, abRect, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectTextOnCurrentArtboardFromLayer(layer.layers[i], abRect, options);
            }
        }

        function collectTextOnAllArtboardsFromLayer(layer, options) {
            if (!layer) return;
            collectTextOnAllArtboardsFromContainer(layer.pageItems, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectTextOnAllArtboardsFromLayer(layer.layers[i], options);
            }
        }

        function collectTextFromDocument(options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextFromLayer(doc.layers[i], options);
            }
        }

        function collectTextOnCurrentArtboardFromDocument(abRect, options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextOnCurrentArtboardFromLayer(doc.layers[i], abRect, options);
            }
        }

        function collectTextOnAllArtboardsFromDocument(options) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextOnAllArtboardsFromLayer(doc.layers[i], options);
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
        function getOrCreateTempLayer() {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === TEMP_LAYER_NAME && doc.layers[i].note === TEMP_LAYER_NOTE) {
                    return { layer: doc.layers[i], created: false };
                }
            }
            var layer = doc.layers.add();
            layer.name = TEMP_LAYER_NAME;
            layer.note = TEMP_LAYER_NOTE;
            return { layer: layer, created: true };
        }

        function clearTempLayer(layer) {
            if (!layer) return;
            try {
                while (layer.pageItems.length > 0) {
                    layer.pageItems[0].remove();
                }
            } catch (e) { }
        }

        function removeTempLayer() {
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                if (doc.layers[i].name === TEMP_LAYER_NAME && doc.layers[i].note === TEMP_LAYER_NOTE) {
                    try {
                        clearTempLayer(doc.layers[i]);
                    } catch (e) { }
                    try {
                        doc.layers[i].remove();
                    } catch (e2) { }
                    break;
                }
            }
        }

        function duplicateSymbolItemToLayer(symbolItem, targetLayer) {
            var duplicateItem = symbolItem.duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING);
            duplicateItem.selected = false;
            return duplicateItem;
        }
        // シンボル内テキスト収集 / Collect text inside symbols
        // =========================================

        function collectTextFramesFromItem(item, result) {
            if (!item) return;

            if (item.typename === 'TextFrame') {
                if (!isEmptyTextFrame(item)) {
                    result.push(item);
                }
                return;
            }

            if (!item.pageItems) return;
            for (var i = 0; i < item.pageItems.length; i++) {
                collectTextFramesFromItem(item.pageItems[i], result);
            }
        }

        function extractTextContentsFromItems(items) {
            var textFrames = [];
            var texts = [];

            for (var i = 0; i < items.length; i++) {
                collectTextFramesFromItem(items[i], textFrames);
            }

            for (var j = 0; j < textFrames.length; j++) {
                texts.push(textFrames[j].contents);
            }

            return texts;
        }

        function removeItems(items) {
            for (var i = items.length - 1; i >= 0; i--) {
                try {
                    if (items[i] && items[i].isValid !== false) {
                        items[i].remove();
                    }
                } catch (e) { }
            }
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
            var processed = {};

            for (var i = 0; i < symbolItems.length; i++) {
                var symbolItem = symbolItems[i];
                if (!symbolItem || symbolItem.isValid === false) continue;

                var symbolName = '';
                try {
                    symbolName = symbolItem.symbol.name;
                } catch (e) { }

                if (!symbolName || processed[symbolName]) continue;
                processed[symbolName] = true;
                siList.push(symbolItem);
            }

            if (siList.length === 0) {
                return results;
            }

            var tempInfo = getOrCreateTempLayer();
            var tempLayer = tempInfo.layer;
            clearTempLayer(tempLayer);

            withSelectionRestored(function () {
                for (var index = siList.length - 1; index >= 0; index--) {
                    var symbolItem = siList[index];
                    if (!symbolItem || symbolItem.isValid === false) continue;

                    var symbolName = '';
                    try {
                        symbolName = symbolItem.symbol.name;
                    } catch (e) { }

                    try {
                        var workingSymbolItem = duplicateSymbolItemToLayer(symbolItem, tempLayer);
                        try {
                            doc.selection = null;
                        } catch (e2) { }
                        workingSymbolItem.selected = true;
                        workingSymbolItem.breakLink();

                        var brokenItems = [];
                        var newSel = [];
                        try {
                            for (var s = 0; s < doc.selection.length; s++) {
                                newSel.push(doc.selection[s]);
                            }
                        } catch (e3) { }

                        for (var m = 0; m < newSel.length; m++) {
                            brokenItems.push(newSel[m]);
                        }

                        if (brokenItems.length > 0) {
                            var texts = extractTextContentsFromItems(brokenItems);
                            for (var t = 0; t < texts.length; t++) {
                                results.push(symbolName + '：' + texts[t].replace(/[\r\n]+/g, ' '));
                            }
                            removeItems(brokenItems);
                        }
                    } catch (e4) {
                    } finally {
                        try {
                            doc.selection = null;
                        } catch (e5) { }
                    }
                }
            });

            clearTempLayer(tempLayer);
            if (tempInfo.created) {
                removeTempLayer();
            }
            try {
                doc.selection = null;
            } catch (e6) { }

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

        function updateSymbolListHeight(itemCount) {
            var rows = itemCount;
            if (rows < SYMBOL_LIST_MIN_ROWS) rows = SYMBOL_LIST_MIN_ROWS;
            if (rows > SYMBOL_LIST_MAX_ROWS) rows = SYMBOL_LIST_MAX_ROWS;
            symbolListBox.preferredSize.height = rows * SYMBOL_LIST_ROW_HEIGHT + SYMBOL_LIST_EXTRA_HEIGHT;
            try {
                dlg.layout.layout(true);
                dlg.layout.resize();
            } catch (e) { }
        }

        function refreshSymbolList(forceRefresh) {
            symbolListBox.removeAll();
            var symTexts = getSymbolTexts(forceRefresh);
            for (var s = 0; s < symTexts.length; s++) {
                symbolListBox.add("item", symTexts[s]);
            }
            updateSymbolListHeight(symTexts.length);
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

            leftCol.add("statictext", undefined, L("textListLabel"));
            var listBox = leftCol.add("listbox", [0, 0, 250, 194], []);
            leftCol.add("statictext", undefined, L("textEditLabel"));
            var editBox = leftCol.add("edittext", [0, 0, 250, 72], "", { multiline: true, scrolling: true });

            /* シンボル内テキスト / Text in symbols */
            leftCol.add("statictext", undefined, L("symbolTextLabel"));
            var symbolListBox = leftCol.add("listbox", [0, 0, 250, SYMBOL_LIST_MIN_ROWS * SYMBOL_LIST_ROW_HEIGHT + SYMBOL_LIST_EXTRA_HEIGHT], []);
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

            var isPreview = rightCol.add("checkbox", undefined, L("preview"));
            isPreview.helpTip = L("previewTip");

            var buttonRow = dlg.add("group");
            buttonRow.orientation = "row";
            buttonRow.alignChildren = ["fill", "center"];

            var leftButtons = buttonRow.add("group");
            leftButtons.orientation = "row";
            leftButtons.alignChildren = ["left", "center"];
            var exportTextBtn = leftButtons.add("button", undefined, L("exportText"));

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
                exportTextBtn: exportTextBtn,
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
        var exportTextBtn = ui.exportTextBtn;
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
            if (isFormat.value && isPreview.value) {
                isPreview.value = false;
                preview();
            }
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
        exportTextBtn.onClick = function () {
            var filePath;
            var file = null;
            var content;
            try {
                content = buildExportText();
                filePath = Folder.desktop.fsName + '/text-' + sanitizeFileName(getDocumentBaseName(doc)) + '-' + getDateStamp() + '.txt';
                file = new File(filePath);
                file.encoding = 'UTF-8';
                file.lineFeed = 'Unix';
                if (!file.open('w')) {
                    throw new Error('open failed: ' + filePath);
                }
                file.write(content);
                file.close();
                alert(L('exportDone') + '\n' + file.fsName);
            } catch (e) {
                try {
                    if (file && file.opened) file.close();
                } catch (closeErr) { }
                alert(L('exportFailed') + '\n' + e);
            }
        };

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

        /* テキスト書き出し用テキスト生成 / Build export text grouped by artboard */
        function buildExportText() {
            var artboardGroups = collectArtboardGroupedExportData();
            var lines = [];
            var i;
            var j;

            for (i = 0; i < artboardGroups.length; i++) {
                if (i > 0) lines.push('');
                lines.push('---' + artboardGroups[i].name + '---');
                for (j = 0; j < artboardGroups[i].texts.length; j++) {
                    lines.push(artboardGroups[i].texts[j]);
                }
            }

            return lines.join('\n');
        }

        function getArtboardDisplayName(index) {
            var name = '';
            var numberText = '';
            try {
                name = doc.artboards[index].name;
            } catch (e) { }
            if (!name) {
                name = (lang === 'ja') ? 'アートボード ' + (index + 1) : 'Artboard ' + (index + 1);
            }
            numberText = (lang === 'ja') ? ('アートボード' + (index + 1)) : ('Artboard ' + (index + 1));
            return numberText + ': ' + name;
        }

        function getItemArtboardIndex(item) {
            for (var i = 0; i < doc.artboards.length; i++) {
                if (isOnArtboard(item, doc.artboards[i].artboardRect)) {
                    return i;
                }
            }
            return -1;
        }

        function collectAllTextFramesFromContainer(items, result) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) continue;

                if (item.typename === 'TextFrame') {
                    if (!isEmptyTextFrame(item)) {
                        result.push(item);
                    }
                } else if (item.typename === 'GroupItem') {
                    collectAllTextFramesFromContainer(item.pageItems, result);
                }
            }
        }

        function collectAllTextFramesFromLayer(layer, result) {
            if (!layer) return;
            collectAllTextFramesFromContainer(layer.pageItems, result);
            for (var i = 0; i < layer.layers.length; i++) {
                collectAllTextFramesFromLayer(layer.layers[i], result);
            }
        }

        function collectAllDocumentTextFrames() {
            var frames = [];
            for (var i = 0; i < doc.layers.length; i++) {
                collectAllTextFramesFromLayer(doc.layers[i], frames);
            }
            return frames;
        }

        function collectSymbolTextsByArtboard() {
            var results = [];
            var symbolItems = doc.symbolItems;
            var processed = {};
            var tempInfo;
            var tempLayer;

            for (var i = 0; i < symbolItems.length; i++) {
                var symbolItem = symbolItems[i];
                var symbolName = '';
                var artboardIndex = -1;
                if (!symbolItem || symbolItem.isValid === false) continue;

                try {
                    symbolName = symbolItem.symbol.name;
                } catch (e) { }
                if (!symbolName || processed[symbolName]) continue;

                artboardIndex = getItemArtboardIndex(symbolItem);
                if (artboardIndex < 0) continue;

                processed[symbolName] = true;
            }

            tempInfo = getOrCreateTempLayer();
            tempLayer = tempInfo.layer;
            clearTempLayer(tempLayer);

            withSelectionRestored(function () {
                for (var j = 0; j < symbolItems.length; j++) {
                    var sourceItem = symbolItems[j];
                    var sourceSymbolName = '';
                    var sourceArtboardIndex = -1;
                    var brokenItems = [];
                    var newSel = [];
                    var texts;
                    var t;

                    if (!sourceItem || sourceItem.isValid === false) continue;

                    try {
                        sourceSymbolName = sourceItem.symbol.name;
                    } catch (e2) { }
                    if (!sourceSymbolName || processed[sourceSymbolName] !== true) continue;

                    sourceArtboardIndex = getItemArtboardIndex(sourceItem);
                    if (sourceArtboardIndex < 0) {
                        processed[sourceSymbolName] = false;
                        continue;
                    }
                    if (processed[sourceSymbolName] === false) continue;

                    processed[sourceSymbolName] = false;

                    try {
                        var workingSymbolItem = duplicateSymbolItemToLayer(sourceItem, tempLayer);
                        try {
                            doc.selection = null;
                        } catch (e3) { }
                        workingSymbolItem.selected = true;
                        workingSymbolItem.breakLink();

                        try {
                            for (var s = 0; s < doc.selection.length; s++) {
                                newSel.push(doc.selection[s]);
                            }
                        } catch (e4) { }

                        for (var m = 0; m < newSel.length; m++) {
                            brokenItems.push(newSel[m]);
                        }

                        if (brokenItems.length > 0) {
                            texts = extractTextContentsFromItems(brokenItems);
                            for (t = 0; t < texts.length; t++) {
                                var openParen = (lang === 'ja') ? '（' : ' (';
                                var closeParen = (lang === 'ja') ? '）' : ')';
                                var prefix = (lang === 'ja') ? 'シンボル：' : 'Symbol: ';
                                results.push({
                                    artboardIndex: sourceArtboardIndex,
                                    text: texts[t] + openParen + prefix + sourceSymbolName + closeParen
                                });
                            }
                            removeItems(brokenItems);
                        }
                    } catch (e5) {
                    } finally {
                        try {
                            doc.selection = null;
                        } catch (e6) { }
                    }
                }
            });

            clearTempLayer(tempLayer);
            if (tempInfo.created) {
                removeTempLayer();
            }
            try {
                doc.selection = null;
            } catch (e7) { }

            return results;
        }

        function collectArtboardGroupedExportData() {
            var groups = [];
            var outsideGroup = {
                name: (lang === 'ja') ? 'アートボード外' : 'Outside Artboards',
                texts: []
            };
            var i;
            var frames = collectAllDocumentTextFrames();
            var symbols = collectSymbolTextsByArtboard();
            var abIndex;

            for (i = 0; i < doc.artboards.length; i++) {
                groups.push({
                    name: getArtboardDisplayName(i),
                    texts: []
                });
            }

            for (i = 0; i < frames.length; i++) {
                abIndex = getItemArtboardIndex(frames[i]);
                if (abIndex >= 0) {
                    groups[abIndex].texts.push(frames[i].contents);
                } else {
                    outsideGroup.texts.push(frames[i].contents);
                }
            }

            for (i = 0; i < symbols.length; i++) {
                abIndex = symbols[i].artboardIndex;
                if (abIndex >= 0) {
                    groups[abIndex].texts.push(symbols[i].text);
                } else {
                    outsideGroup.texts.push(symbols[i].text);
                }
            }

            if (outsideGroup.texts.length > 0) {
                groups.push(outsideGroup);
            }
            return groups;
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
    }
}

main();
