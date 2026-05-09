#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.4.14";

/*

TextScopeEdit.jsx

概要
選択した条件に応じてドキュメント内のテキストを収集し、一覧表示できます。
対象テキスト（アートボード）では、現在のアートボード内 / すべてのアートボード内 / ドキュメント全体を切り替えできます。
対象テキスト（レイヤー）では、//ではじめるレイヤーを含めるかどうか、ロックされたテキスト、非表示のテキストを対象に含めるかどうかを切り替えできます。
同じ内容のテキストも個別に一覧表示します。［重複を削除］をオンにすると、同じ内容のテキストを一覧では1件だけ表示します。
グループ内およびサブレイヤー内のテキストも対象です。実質的に空のテキストオブジェクトは無視します。
シンボル内のテキストはテキスト一覧に統合して表示します。シンボル内テキストの読み取り結果はダイアログ表示中にキャッシュされますが、対象アートボードや対象レイヤー条件が変わると現在の条件に合わせて再取得します。
シンボル内テキストの読み取りは、同名シンボルを1回だけ対象にします。対象シンボルの代表インスタンスを一時レイヤーへ複製し、複製側だけを breakLink して展開後のテキストを収集します。読み取りに使った複製オブジェクトは削除し、元のカンバス上のシンボルには触れません。一時レイヤーは必要時のみ作成し、このスクリプトが作成した場合のみ削除します。
ダイアログ左側には「テキスト一覧」を複数行のテキスト欄で表示します。通常テキストとシンボル内テキストは同じ一覧にまとめて表示します。［テキストをコピー］で現在のテキスト一覧を選択してコピーできます。
［テキスト書き出し］で、現在の一覧条件に応じた通常テキストとシンボル内テキストをアートボードごとにまとめてデスクトップへテキストファイルとして書き出します。ファイル名には yyyymmdd-hhmmss 形式の日時を付与します。現在のアートボード内 / すべてのアートボード内 / ドキュメント全体を対象に、//ではじめるレイヤーを含めるかどうか、ロックされたテキスト、非表示のテキストのチェック状態も書き出しに反映されます。現在のアートボード内では現在のアートボードだけを書き出し、ほかのアートボード名は出力しません。すべてのアートボード内 / ドキュメント全体では、各アートボードは ---アートボード番号: アートボード名--- / アートボード外は ---アートボード外--- の見出しで区切って出力します。［テキスト書き出し］で書き出して閉じます。［キャンセル］でダイアログを閉じます。

Overview
Collect text in the document based on the selected conditions and show it in a list.
In Text Scope (Artboards), you can switch between the current artboard, all artboards, and the entire document.
In Text Scope (Layers), you can choose whether to include layers starting with //, and whether locked text and hidden text are included.
Text frames with the same content are also listed individually. When [Remove Duplicates] is enabled, only one item per identical text content is shown in the list.
Text inside groups and sublayers is supported. Effectively empty text objects are ignored.
Text inside symbols is merged into the main text list. Symbol text results are cached while the dialog is open, and normal list updates do not reload them.
To read text inside symbols, the script targets each symbol name only once. It duplicates one representative instance of the target symbol to a temporary layer, break-links only the duplicate, and collects the resulting text. The temporary duplicate objects are then removed, and the original symbols on the canvas remain untouched. The temporary layer is created only when needed and is removed only if this script created it.
The left side of the dialog shows Text List in a multi-line text field. Regular text and text found inside symbols are displayed together in the same list. [Copy Text] selects the current text list and copies it to the clipboard.
[Export Text] writes the regular text and the text found inside symbols that match the current list conditions to a text file on the Desktop. The filename includes a timestamp in yyyymmdd-hhmmss format. The export reflects the current artboard scope setting (Current Artboard Only, All Artboards Only, or Include Entire Document), and it also reflects the current checkbox states for including layers that start with //, locked text, and hidden text. In Current Artboard Only, only the current artboard is written and no other artboard headings are output. In All Artboards Only or Include Entire Document, each artboard is written under a ---Artboard Number: Artboard Name--- heading, and items outside all artboards are written under ---Outside Artboards---. Text found inside symbols is also grouped under the artboard where that symbol instance is placed. Pressing [Export Text] exports and closes the dialog. [Cancel] closes the dialog.

更新日 / Updated: 2026-04-03

紹介記事（note）
https://note.com/dtp_tranist/n/nb845889dd553

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
    dialogTitle: { ja: "テキスト一覧と書き出し", en: "Text Collector & Exporter" },
    panelTargetText: { ja: "対象テキスト（アートボード）", en: "Text Scope (Artboards)" },
    panelLayerText: { ja: "対象テキスト（レイヤー）", en: "Text Scope (Layers)" },
    rbCurrentArtboard: { ja: "現在のアートボード内", en: "Current Artboard Only" },
    rbAllArtboards: { ja: "すべてのアートボード内", en: "All Artboards Only" },
    cbOutside: { ja: "ドキュメント全体を対象に", en: "Include Entire Document" },
    cbRemoveDuplicates: { ja: "重複を削除", en: "Remove Duplicates" },
    cbSkipComment: { ja: "//ではじめるレイヤーも含む", en: "Include Layers That Start with //" },
    cbIncludeLocked: { ja: "ロックされたテキスト", en: "Locked Text" },
    cbIncludeHidden: { ja: "非表示のテキスト", en: "Hidden Text" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    exportAndClose: { ja: "テキスト書き出し", en: "Export Text" },
    copyText: { ja: "テキストをコピー", en: "Copy Text" },
    exportDone: { ja: "テキストを書き出しました", en: "Text exported" },
    exportFailed: { ja: "テキストを書き出せませんでした", en: "Failed to export text" },
    noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
    textListLabel: { ja: "テキスト一覧", en: "Text List" },
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
    try {

        /* テキストフレームの参照を保持する配列 / Store references to text frames */
        var textFrameList = [];

        var symbolTextCache = null;
        var symbolTextCacheKey = null;
        var TEMP_LAYER_NAME = '__TextScopeEdit_temp_read__';
        var TEMP_LAYER_NOTE = '__TextScopeEdit_temp_read__';

        /* 一覧表示用のラベルを生成（先頭部分を表示） / Build labels for the list display using the leading part of the text */
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
            return now.getFullYear()
                + zeroPad2(now.getMonth() + 1)
                + zeroPad2(now.getDate())
                + '-'
                + zeroPad2(now.getHours())
                + zeroPad2(now.getMinutes())
                + zeroPad2(now.getSeconds());
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

        function getExportOptions() {
            var mode;
            if (cbOutside.value) {
                mode = "all";
            } else if (rbAll.value) {
                mode = "allArtboards";
            } else {
                mode = "current";
            }
            return {
                mode: mode,
                includeComment: cbSkipComment.value,
                includeLocked: cbIncludeLocked.value,
                includeHidden: cbIncludeHidden.value
            };
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

        function withDocumentSelectionCleared(fn) {
            var prevSelection = [];
            try {
                for (var i = 0; i < doc.selection.length; i++) {
                    prevSelection.push(doc.selection[i]);
                }
            } catch (e) { }

            try {
                try {
                    doc.selection = null;
                } catch (e2) { }
                fn();
            } finally {
                try {
                    doc.selection = null;
                } catch (e3) { }
                for (var j = 0; j < prevSelection.length; j++) {
                    try {
                        prevSelection[j].selected = true;
                    } catch (e4) { }
                }
            }
        }

        function collectSymbolTexts(mode) {
            var results = [];
            var symbolItems = doc.symbolItems;
            var siList = [];
            var processed = {};

            for (var i = 0; i < symbolItems.length; i++) {
                var symbolItem = symbolItems[i];
                var symbolName = '';
                if (!symbolItem || symbolItem.isValid === false) continue;

                try {
                    symbolName = symbolItem.symbol.name;
                } catch (e) { }
                if (!symbolName || processed[symbolName]) continue;

                var symbolArtboardIndex = getItemArtboardIndex(symbolItem);
                if (mode === "current" && symbolArtboardIndex !== doc.artboards.getActiveArtboardIndex()) continue;
                if (mode === "allArtboards" && symbolArtboardIndex < 0) continue;

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
                                results.push(texts[t].replace(/[\r\n]+/g, ' '));
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
        // テキスト一覧を更新 / Update text list
        // =========================================

        function updateList() {
            var mode;
            var lines = [];
            if (cbOutside.value) {
                mode = "all";
            } else if (rbAll.value) {
                mode = "allArtboards";
            } else {
                mode = "current";
            }
            gatherFrames(mode);
            var symTexts = getSymbolTexts(false, mode);
            for (var s = 0; s < symTexts.length; s++) {
                textFrameList.push({ contents: symTexts[s] });
            }
            if (cbRemoveDuplicates.value) {
                textFrameList = removeDuplicateTextEntries(textFrameList);
            }
            for (var i = 0; i < textFrameList.length; i++) {
                lines.push(makeLabel(textFrameList[i].contents));
            }
            listBox.text = lines.join("\n");
        }

        function getSymbolTexts(forceRefresh, mode) {
            var cacheKey = [
                mode,
                cbSkipComment.value ? '1' : '0',
                cbIncludeLocked.value ? '1' : '0',
                cbIncludeHidden.value ? '1' : '0'
            ].join('|');

            if (!forceRefresh && symbolTextCache !== null && symbolTextCacheKey === cacheKey) {
                return symbolTextCache;
            }

            symbolTextCache = collectSymbolTexts(mode);
            symbolTextCacheKey = cacheKey;
            return symbolTextCache;
        }

        function removeDuplicateTextEntries(items) {
            var uniqueItems = [];
            var duplicateMap = {};
            for (var i = 0; i < items.length; i++) {
                var text = items[i] && items[i].contents ? items[i].contents : "";
                if (duplicateMap[text] === undefined) {
                    duplicateMap[text] = true;
                    uniqueItems.push(items[i]);
                }
            }
            return uniqueItems;
        }


        // =========================================
        // ダイアログ / Dialog
        // =========================================

        function buildDialogUI() {
            var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];

            /* 2カラム: 左=listbox群、右=オプション、下=ボタンエリア / Two columns: left = list boxes, right = options, bottom = button area */
            var mainGroup = dlg.add("group");
            mainGroup.orientation = "row";
            mainGroup.alignChildren = ["fill", "fill"];

            /* 左カラム / Left column */
            var leftCol = mainGroup.add("group");
            leftCol.orientation = "column";
            leftCol.alignChildren = ["fill", "top"];

            leftCol.add("statictext", undefined, L("textListLabel"));
            var listBox = leftCol.add("edittext", [0, 0, 250, 284], "", { multiline: true, scrolling: true, readonly: true });

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

            var dupGroup = rightCol.add("group");
            dupGroup.orientation = "row";
            dupGroup.alignChildren = ["center", "center"];
            dupGroup.alignment = ["center", "top"];
            var cbRemoveDuplicates = dupGroup.add("checkbox", undefined, L("cbRemoveDuplicates"));
            cbRemoveDuplicates.value = false;

            var buttonRow = dlg.add("group");
            buttonRow.orientation = "row";
            buttonRow.alignChildren = ["fill", "center"];

            var leftButtons = buttonRow.add("group");
            leftButtons.orientation = "row";
            leftButtons.alignChildren = ["left", "center"];
            var cancelBtn = leftButtons.add("button", undefined, L("cancel"), { name: "cancel" });

            var spacer = buttonRow.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var rightButtons = buttonRow.add("group");
            rightButtons.orientation = "row";
            rightButtons.alignChildren = ["right", "center"];
            var copyTextBtn = rightButtons.add("button", undefined, L("copyText"));
            var exportTextBtn = rightButtons.add("button", undefined, L("exportAndClose"), { name: "ok" });
            dlg.defaultElement = exportTextBtn;

            return {
                dlg: dlg,
                rbArtboard: rbArtboard,
                rbAll: rbAll,
                cbOutside: cbOutside,
                cbRemoveDuplicates: cbRemoveDuplicates,
                cbSkipComment: cbSkipComment,
                cbIncludeLocked: cbIncludeLocked,
                cbIncludeHidden: cbIncludeHidden,
                listBox: listBox,
                cancelBtn: cancelBtn,
                copyTextBtn: copyTextBtn,
                exportTextBtn: exportTextBtn
            };
        }

        var ui = buildDialogUI();
        var dlg = ui.dlg;
        var rbArtboard = ui.rbArtboard;
        var rbAll = ui.rbAll;
        var cbOutside = ui.cbOutside;
        var cbRemoveDuplicates = ui.cbRemoveDuplicates;
        var cbSkipComment = ui.cbSkipComment;
        var cbIncludeLocked = ui.cbIncludeLocked;
        var cbIncludeHidden = ui.cbIncludeHidden;
        var listBox = ui.listBox;
        var cancelBtn = ui.cancelBtn;
        var copyTextBtn = ui.copyTextBtn;
        var exportTextBtn = ui.exportTextBtn;

        // =========================================
        // イベントハンドラ / Event handlers
        // =========================================
        copyTextBtn.onClick = function () {
            try {
                withDocumentSelectionCleared(function () {
                    listBox.active = true;
                    listBox.textselection = listBox.text;
                    app.executeMenuCommand("copy");
                });
            } catch (e) {
                alert(e);
            }
        };

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
                dlg.close();
            } catch (e) {
                try {
                    if (file && file.opened) file.close();
                } catch (closeErr) { }
                alert(L('exportFailed') + '\n' + e);
            }
        };

        /* テキスト書き出し用テキスト生成 / Build export text grouped by artboard */
        function buildExportText() {
            var artboardGroups = collectArtboardGroupedExportData(getExportOptions());
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

        function collectAllTextFramesOnArtboardFromContainer(items, abRect, result, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) continue;
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === 'TextFrame') {
                    if (!isEmptyTextFrame(item) && isOnArtboard(item, abRect)) {
                        result.push(item);
                    }
                } else if (item.typename === 'GroupItem') {
                    collectAllTextFramesOnArtboardFromContainer(item.pageItems, abRect, result, options);
                }
            }
        }

        function collectAllTextFramesOnAllArtboardsFromContainer(items, result, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) continue;
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === 'TextFrame') {
                    if (!isEmptyTextFrame(item) && isOnAnyArtboard(item)) {
                        result.push(item);
                    }
                } else if (item.typename === 'GroupItem') {
                    collectAllTextFramesOnAllArtboardsFromContainer(item.pageItems, result, options);
                }
            }
        }

        function collectAllTextFramesFromContainer(items, result, options) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) continue;
                if (!options.includeComment && isCommentLayer(item)) continue;
                if (!isCollectable(item, options)) continue;

                if (item.typename === 'TextFrame') {
                    if (!isEmptyTextFrame(item)) {
                        result.push(item);
                    }
                } else if (item.typename === 'GroupItem') {
                    collectAllTextFramesFromContainer(item.pageItems, result, options);
                }
            }
        }

        function collectAllTextFramesFromLayer(layer, result, options) {
            if (!layer) return;
            collectAllTextFramesFromContainer(layer.pageItems, result, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectAllTextFramesFromLayer(layer.layers[i], result, options);
            }
        }

        function collectAllTextFramesOnCurrentArtboardFromLayer(layer, abRect, result, options) {
            if (!layer) return;
            collectAllTextFramesOnArtboardFromContainer(layer.pageItems, abRect, result, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectAllTextFramesOnCurrentArtboardFromLayer(layer.layers[i], abRect, result, options);
            }
        }

        function collectAllTextFramesOnAllArtboardsFromLayer(layer, result, options) {
            if (!layer) return;
            collectAllTextFramesOnAllArtboardsFromContainer(layer.pageItems, result, options);
            for (var i = 0; i < layer.layers.length; i++) {
                collectAllTextFramesOnAllArtboardsFromLayer(layer.layers[i], result, options);
            }
        }

        function collectAllDocumentTextFrames(options) {
            var frames = [];
            var i;
            if (options.mode === 'all') {
                for (i = 0; i < doc.layers.length; i++) {
                    collectAllTextFramesFromLayer(doc.layers[i], frames, options);
                }
            } else if (options.mode === 'allArtboards') {
                for (i = 0; i < doc.layers.length; i++) {
                    collectAllTextFramesOnAllArtboardsFromLayer(doc.layers[i], frames, options);
                }
            } else {
                var abIndex = doc.artboards.getActiveArtboardIndex();
                var abRect = doc.artboards[abIndex].artboardRect;
                for (i = 0; i < doc.layers.length; i++) {
                    collectAllTextFramesOnCurrentArtboardFromLayer(doc.layers[i], abRect, frames, options);
                }
            }
            return frames;
        }

        function collectSymbolTextsByArtboard(options) {
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
                if (!options.includeComment && isCommentLayer(symbolItem)) continue;
                if (!isCollectable(symbolItem, options)) continue;

                try {
                    symbolName = symbolItem.symbol.name;
                } catch (e) { }
                if (!symbolName || processed[symbolName]) continue;

                artboardIndex = getItemArtboardIndex(symbolItem);
                if (options.mode === 'current') {
                    if (artboardIndex !== doc.artboards.getActiveArtboardIndex()) continue;
                } else if (options.mode === 'allArtboards') {
                    if (artboardIndex < 0) continue;
                }

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
                    if (!options.includeComment && isCommentLayer(sourceItem)) continue;
                    if (!isCollectable(sourceItem, options)) continue;

                    try {
                        sourceSymbolName = sourceItem.symbol.name;
                    } catch (e2) { }
                    if (!sourceSymbolName || processed[sourceSymbolName] !== true) continue;

                    sourceArtboardIndex = getItemArtboardIndex(sourceItem);
                    if (options.mode === 'current') {
                        if (sourceArtboardIndex !== doc.artboards.getActiveArtboardIndex()) {
                            processed[sourceSymbolName] = false;
                            continue;
                        }
                    } else if (options.mode === 'allArtboards') {
                        if (sourceArtboardIndex < 0) {
                            processed[sourceSymbolName] = false;
                            continue;
                        }
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
                                results.push({
                                    artboardIndex: sourceArtboardIndex,
                                    text: texts[t]
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

        function collectArtboardGroupedExportData(options) {
            var groups = [];
            var outsideGroup = {
                name: (lang === 'ja') ? 'アートボード外' : 'Outside Artboards',
                texts: []
            };
            var i;
            var frames = collectAllDocumentTextFrames(options);
            var symbols = collectSymbolTextsByArtboard(options);
            var abIndex;
            var activeIndex = doc.artboards.getActiveArtboardIndex();

            if (options.mode === 'current') {
                groups.push({
                    name: getArtboardDisplayName(activeIndex),
                    texts: []
                });
            } else {
                for (i = 0; i < doc.artboards.length; i++) {
                    groups.push({
                        name: getArtboardDisplayName(i),
                        texts: []
                    });
                }
            }

            for (i = 0; i < frames.length; i++) {
                abIndex = getItemArtboardIndex(frames[i]);
                if (options.mode === 'current') {
                    if (abIndex === activeIndex) {
                        groups[0].texts.push(frames[i].contents);
                    }
                } else if (abIndex >= 0) {
                    groups[abIndex].texts.push(frames[i].contents);
                } else {
                    outsideGroup.texts.push(frames[i].contents);
                }
            }

            for (i = 0; i < symbols.length; i++) {
                abIndex = symbols[i].artboardIndex;
                if (options.mode === 'current') {
                    if (abIndex === activeIndex) {
                        groups[0].texts.push(symbols[i].text);
                    }
                } else if (abIndex >= 0) {
                    groups[abIndex].texts.push(symbols[i].text);
                } else {
                    outsideGroup.texts.push(symbols[i].text);
                }
            }

            if (options.mode !== 'current' && outsideGroup.texts.length > 0) {
                groups.push(outsideGroup);
            }
            return groups;
        }

        /* 初回収集 / Initial collection */
        updateList();

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

        cbOutside.onClick = function () { updateList(); };
        cbRemoveDuplicates.onClick = function () { updateList(); };
        cbSkipComment.onClick = function () { updateList(); };
        cbIncludeLocked.onClick = function () { updateList(); };
        cbIncludeHidden.onClick = function () { updateList(); };

        /* キャンセルボタンで閉じる / Close the dialog when Cancel is pressed */
        cancelBtn.onClick = function () {
            dlg.close();
        };


        dlg.onClose = function () {
        };

        dlg.show();
    } finally {
    }
}

main();