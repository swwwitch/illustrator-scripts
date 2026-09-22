#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストを収集して一覧表示し、ファイルへ書き出します。
対象はアートボード単位やレイヤー単位で絞り込め、重複をまとめることもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextExport.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb845889dd553

### Overview

Collects the text in the document, lists it, and exports it to a file.
The scope can be narrowed by artboard or by layer, and duplicates can be folded together.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextExport.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextExport";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.15";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextExport.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextExport.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb845889dd553"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 / panel margins */
    var TEXT_LIST_BOUNDS = [0, 0, 250, 284];    /* テキスト一覧の大きさ / text list bounds */
    var LIST_LINE_MAX_CHARS = 40;               /* 一覧の1行に出す最大文字数 / max characters per list line */

    // =========================================
    // 一時レイヤー / Temporary layer
    // =========================================
    /* シンボルの中身を読むための一時レイヤー（TextScopeEdit と同じ名前） / Temporary layer for reading symbols */
    var TEMP_LAYER_NAME = '__TextScopeEdit_temp_read__';
    var TEMP_LAYER_NOTE = '__TextScopeEdit_temp_read__';

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "テキスト一覧と書き出し", en: "Text Collector & Exporter" }
        },
        panel: {
            targetArtboards: { ja: "対象テキスト（アートボード）", en: "Text Scope (Artboards)" },
            targetLayers: { ja: "対象テキスト（レイヤー）", en: "Text Scope (Layers)" }
        },
        fieldLabel: {
            textList: { ja: "テキスト一覧", en: "Text List" }
        },
        radio: {
            currentArtboard: { ja: "現在のアートボード内", en: "Current Artboard Only" },
            allArtboards: { ja: "すべてのアートボード内", en: "All Artboards Only" }
        },
        checkbox: {
            includeOutside: { ja: "ドキュメント全体を対象に", en: "Include Entire Document" },
            removeDuplicates: { ja: "重複を削除", en: "Remove Duplicates" },
            includeCommentLayers: { ja: "//ではじめるレイヤーも含む", en: "Include Layers That Start with //" },
            includeLocked: { ja: "ロックされたテキスト", en: "Locked Text" },
            includeHidden: { ja: "非表示のテキスト", en: "Hidden Text" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            exportText: { ja: "テキスト書き出し", en: "Export Text" },
            copyText: { ja: "テキストをコピー", en: "Copy Text" }
        },
        alert: {
            exportDone: { ja: "テキストを書き出しました", en: "Text exported" },
            exportFailed: { ja: "テキストを書き出せませんでした", en: "Failed to export text" },
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" }
        },
        tooltip: {
            textList: {
                ja: "集めたテキストの一覧です。ここで編集した内容が書き出し・コピーの対象になります。",
                en: "The collected text. What you edit here is what gets exported or copied."
            },
            currentArtboard: {
                ja: "いま表示しているアートボードの中のテキストだけを集めます。",
                en: "Collects only the text on the artboard currently in view."
            },
            allArtboards: { ja: "すべてのアートボードの中のテキストを集めます。", en: "Collects the text on every artboard." },
            includeOutside: {
                ja: "アートボードの外に置かれたテキストも含めます。",
                en: "Also includes text placed outside the artboards."
            },
            includeCommentLayers: {
                ja: "レイヤー名が // ではじまるレイヤーのテキストも集めます。",
                en: "Also collects text on layers whose name starts with //."
            },
            includeLocked: { ja: "ロックされたテキストも集めます。", en: "Also collects locked text." },
            includeHidden: { ja: "非表示のテキストも集めます。", en: "Also collects hidden text." },
            removeDuplicates: { ja: "同じ内容のテキストを1つにまとめます。", en: "Keeps only one copy of identical text." },
            exportText: {
                ja: "アートボードごとに見出しを付けたテキストファイルをデスクトップに保存し、ダイアログを閉じます。一覧の省略表示と［重複を削除］は反映されません。",
                en: "Saves the text to a file on the Desktop, with a heading for each artboard, and closes the dialog. The shortened list lines and Remove Duplicates do not apply."
            },
            copyText: { ja: "一覧の内容をクリップボードへコピーします。", en: "Copies the list to the clipboard." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文言を取り出す
     * @param {string} labelPath - "dialog.title" のようなパス
     * @returns {string} 表示言語の文言（見つからないときはパスそのもの）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 文字列・ファイル名 / Strings and file names
    // =========================================

    /**
     * 一覧に出す1行を作る（改行をスペースにし、長いものは先頭だけ）
     * @param {string} sourceText - テキストの内容
     * @returns {string} 一覧用の1行
     */
    function makeListLine(sourceText) {
        var lineText = sourceText.replace(/[\r\n]+/g, " ");
        if (lineText.length > LIST_LINE_MAX_CHARS) lineText = lineText.substring(0, LIST_LINE_MAX_CHARS) + "…";
        return lineText;
    }

    /**
     * 2桁になるよう 0 を補う
     * @param {number} numberValue - 対象の数値
     * @returns {string} 2桁の文字列
     */
    function padTwoDigits(numberValue) {
        return (numberValue < 10 ? '0' : '') + numberValue;
    }

    /**
     * ファイル名用の日時 YYYYMMDD-HHMMSS を返す
     * @returns {string} 日時の文字列
     */
    function getDateStamp() {
        var currentDate = new Date();
        return currentDate.getFullYear()
            + padTwoDigits(currentDate.getMonth() + 1)
            + padTwoDigits(currentDate.getDate())
            + '-'
            + padTwoDigits(currentDate.getHours())
            + padTwoDigits(currentDate.getMinutes())
            + padTwoDigits(currentDate.getSeconds());
    }

    /**
     * ドキュメント名から拡張子を除いた名前を返す
     * @param {Document} documentRef - 対象のドキュメント
     * @returns {string} 拡張子なしの名前（名前が無ければ "untitled"）
     */
    function getDocumentBaseName(documentRef) {
        var docName = documentRef && documentRef.name ? documentRef.name : 'untitled';
        return docName.replace(/\.[^\.]+$/, '');
    }

    /**
     * ファイル名に使えない文字を "_" に置き換える
     * @param {string} fileName - 元の名前
     * @returns {string} 置き換えた名前
     */
    function sanitizeFileName(fileName) {
        return fileName.replace(/[\\\/\:\*\?\"\<\>\|]+/g, '_');
    }

    /**
     * 重複した内容のエントリーを除く（最初の1つを残す）
     * @param {Object[]} textEntries - contents を持つエントリーの配列
     * @returns {Object[]} 重複を除いた配列
     */
    function removeDuplicateTextEntries(textEntries) {
        var uniqueEntries = [];
        var seenTexts = {};
        for (var i = 0; i < textEntries.length; i++) {
            var entryText = textEntries[i] && textEntries[i].contents ? textEntries[i].contents : "";
            if (seenTexts[entryText] === undefined) {
                seenTexts[entryText] = true;
                uniqueEntries.push(textEntries[i]);
            }
        }
        return uniqueEntries;
    }

    // =========================================
    // 対象の判定 / Target checks
    // =========================================

    /* アクティブドキュメント（main() で設定） / Active document, set in main() */
    var doc = null;

    /**
     * 実質的に空のテキストか判定する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {boolean} 空（改行・空白だけ）なら true
     */
    function isEmptyTextFrame(textFrame) {
        /* 読み取れないものは空として扱う / treat unreadable frames as empty */
        try {
            if (!textFrame || textFrame.typename !== "TextFrame") return true;
            var flatText = textFrame.contents;
            if (!flatText) return true;
            flatText = flatText.replace(/[\r\n\x03]/g, "").replace(/\s+/g, "");
            return flatText.length === 0;
        } catch (e) {
            return true;
        }
    }

    /**
     * レイヤー名が // ではじまるか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} // ではじまるレイヤーにあれば true
     */
    function isCommentLayer(pageItem) {
        /* layer を持たないアイテムがある / some items have no layer */
        try {
            return pageItem.layer.name.indexOf("//") === 0;
        } catch (err) {
            return false;
        }
    }

    /**
     * 非表示・ロックの設定に照らして集めてよいか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Object} collectOptions - includeLocked / includeHidden を持つ設定
     * @returns {boolean} 集めてよければ true
     */
    function isCollectable(pageItem, collectOptions) {
        /* 状態を読めないアイテムは対象外 / skip items whose state cannot be read */
        try {
            if (!pageItem) return false;
            if (!collectOptions.includeHidden && pageItem.hidden) return false;
            if (!collectOptions.includeLocked && pageItem.locked) return false;
            if (pageItem.layer) {
                if (!collectOptions.includeLocked && pageItem.layer.locked) return false;
                if (!collectOptions.includeHidden && pageItem.layer.visible === false) return false;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アートボードと重なっているか判定する（一部でも重なれば true）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {number[]} artboardRect - [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function isOnArtboard(pageItem, artboardRect) {
        var itemBounds = pageItem.geometricBounds;
        return (itemBounds[2] > artboardRect[0] && itemBounds[0] < artboardRect[2] &&
            itemBounds[1] > artboardRect[3] && itemBounds[3] < artboardRect[1]);
    }

    /**
     * 最初に重なるアートボードの番号を返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number} アートボードの番号（どれにも重ならなければ -1）
     */
    function getItemArtboardIndex(pageItem) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isOnArtboard(pageItem, doc.artboards[i].artboardRect)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * いずれかのアートボードに重なっているか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} 重なっていれば true
     */
    function isOnAnyArtboard(pageItem) {
        return getItemArtboardIndex(pageItem) >= 0;
    }

    /**
     * アートボードの番号が収集範囲に入るか判定する
     * @param {number} artboardIndex - アートボードの番号（-1 はアートボード外）
     * @param {string} collectMode - "all" / "allArtboards" / "current"
     * @returns {boolean} 範囲に入れば true
     */
    function isInArtboardScope(artboardIndex, collectMode) {
        if (collectMode === "current") return artboardIndex === doc.artboards.getActiveArtboardIndex();
        if (collectMode === "allArtboards") return artboardIndex >= 0;
        return true;
    }

    // =========================================
    // テキストフレームの収集 / Collect text frames
    // =========================================

    /**
     * コンテナ内のテキストフレームを再帰的に集める（グループの中もたどる）
     * @param {PageItems} pageItems - 探すアイテムの集まり
     * @param {TextFrame[]} foundFrames - 見つけたテキストを入れる配列
     * @param {function|null} isInScope - 範囲の判定（null なら範囲を問わない）
     * @param {Object} collectOptions - 収集の設定
     * @returns {void}
     */
    function collectTextFramesFromContainer(pageItems, foundFrames, isInScope, collectOptions) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!pageItem) continue;
            if (!collectOptions.includeComment && isCommentLayer(pageItem)) continue;
            if (!isCollectable(pageItem, collectOptions)) continue;

            if (pageItem.typename === 'TextFrame') {
                if (!isEmptyTextFrame(pageItem) && (!isInScope || isInScope(pageItem))) {
                    foundFrames.push(pageItem);
                }
            } else if (pageItem.typename === 'GroupItem') {
                collectTextFramesFromContainer(pageItem.pageItems, foundFrames, isInScope, collectOptions);
            }
        }
    }

    /**
     * レイヤーとそのサブレイヤーからテキストフレームを集める
     * @param {Layer} targetLayer - 対象のレイヤー
     * @param {TextFrame[]} foundFrames - 見つけたテキストを入れる配列
     * @param {function|null} isInScope - 範囲の判定（null なら範囲を問わない）
     * @param {Object} collectOptions - 収集の設定
     * @returns {void}
     */
    function collectTextFramesFromLayer(targetLayer, foundFrames, isInScope, collectOptions) {
        if (!targetLayer) return;
        collectTextFramesFromContainer(targetLayer.pageItems, foundFrames, isInScope, collectOptions);
        for (var i = 0; i < targetLayer.layers.length; i++) {
            collectTextFramesFromLayer(targetLayer.layers[i], foundFrames, isInScope, collectOptions);
        }
    }

    /**
     * 設定の範囲にあるテキストフレームをドキュメント全体から集める
     * @param {Object} collectOptions - mode / includeComment / includeLocked / includeHidden を持つ設定
     * @returns {TextFrame[]} 見つけたテキストフレーム
     */
    function collectDocumentTextFrames(collectOptions) {
        var foundFrames = [];
        var isInScope = null;
        if (collectOptions.mode === 'allArtboards') {
            isInScope = isOnAnyArtboard;
        } else if (collectOptions.mode !== 'all') {
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            isInScope = function (textFrame) {
                return isOnArtboard(textFrame, artboardRect);
            };
        }
        for (var i = 0; i < doc.layers.length; i++) {
            collectTextFramesFromLayer(doc.layers[i], foundFrames, isInScope, collectOptions);
        }
        return foundFrames;
    }

    // =========================================
    // 選択の待避 / Selection helpers
    // =========================================

    /**
     * 選択を解除する
     * @returns {void}
     */
    function clearDocumentSelection() {
        /* テキスト編集中などは解除できないことがある / may fail while editing text */
        try {
            doc.selection = null;
        } catch (e) { }
    }

    /**
     * いまの選択を配列に写して返す
     * @returns {PageItem[]} 選択されているアイテム
     */
    function getSelectedItems() {
        var selectedItems = [];
        /* 選択が読めないときは空とみなす / an unreadable selection counts as empty */
        try {
            for (var i = 0; i < doc.selection.length; i++) {
                selectedItems.push(doc.selection[i]);
            }
        } catch (e) { }
        return selectedItems;
    }

    /**
     * 選択を控えて処理を実行し、終わったら選択を元に戻す
     * @param {function} bodyFn - 実行する処理
     * @param {boolean} clearFirst - 実行前に選択を解除するなら true
     * @returns {void}
     */
    function withSelectionSaved(bodyFn, clearFirst) {
        var savedSelection = getSelectedItems();
        try {
            if (clearFirst) clearDocumentSelection();
            bodyFn();
        } finally {
            clearDocumentSelection();
            for (var i = 0; i < savedSelection.length; i++) {
                /* 消えたアイテムやロックされたアイテムは選び直せない / removed or locked items cannot be reselected */
                try {
                    savedSelection[i].selected = true;
                } catch (e) { }
            }
        }
    }

    // =========================================
    // シンボル内テキスト収集 / Collect text inside symbols
    // =========================================

    /**
     * 一時レイヤーを探し、無ければ作る
     * @returns {{layer: Layer, created: boolean}} 一時レイヤーと、新しく作ったかどうか
     */
    function getOrCreateTempLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === TEMP_LAYER_NAME && doc.layers[i].note === TEMP_LAYER_NOTE) {
                return { layer: doc.layers[i], created: false };
            }
        }
        var tempLayer = doc.layers.add();
        tempLayer.name = TEMP_LAYER_NAME;
        tempLayer.note = TEMP_LAYER_NOTE;
        return { layer: tempLayer, created: true };
    }

    /**
     * 一時レイヤーの中身を消す
     * @param {Layer} tempLayer - 一時レイヤー
     * @returns {void}
     */
    function clearTempLayer(tempLayer) {
        if (!tempLayer) return;
        /* ロックされたアイテムなどは消せない / locked items cannot be removed */
        try {
            while (tempLayer.pageItems.length > 0) {
                tempLayer.pageItems[0].remove();
            }
        } catch (e) { }
    }

    /**
     * 一時レイヤーを削除する
     * @returns {void}
     */
    function removeTempLayer() {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === TEMP_LAYER_NAME && doc.layers[i].note === TEMP_LAYER_NOTE) {
                clearTempLayer(doc.layers[i]);
                /* ロックされたレイヤーなどは削除できない / a locked layer cannot be removed */
                try {
                    doc.layers[i].remove();
                } catch (e) { }
                break;
            }
        }
    }

    /**
     * 一時レイヤーを用意して処理を実行し、終わったら片付ける
     * @param {function} bodyFn - 一時レイヤーを受け取って実行する処理
     * @returns {void}
     */
    function withTempSymbolLayer(bodyFn) {
        var tempInfo = getOrCreateTempLayer();
        var tempLayer = tempInfo.layer;
        clearTempLayer(tempLayer);

        withSelectionSaved(function () {
            bodyFn(tempLayer);
        }, false);

        clearTempLayer(tempLayer);
        if (tempInfo.created) {
            removeTempLayer();
        }
        clearDocumentSelection();
    }

    /**
     * シンボル名を返す
     * @param {SymbolItem} symbolItem - 対象のシンボルインスタンス
     * @returns {string} シンボル名（読めなければ空文字）
     */
    function getSymbolName(symbolItem) {
        /* 定義を失ったシンボルインスタンスがある / an instance may have lost its symbol */
        try {
            return symbolItem.symbol.name;
        } catch (e) {
            return '';
        }
    }

    /**
     * アイテムの中のテキストフレームを再帰的に集める
     * @param {PageItem} pageItem - 対象のアイテム
     * @param {TextFrame[]} foundFrames - 見つけたテキストを入れる配列
     * @returns {void}
     */
    function collectTextFramesFromItem(pageItem, foundFrames) {
        if (!pageItem) return;

        if (pageItem.typename === 'TextFrame') {
            if (!isEmptyTextFrame(pageItem)) {
                foundFrames.push(pageItem);
            }
            return;
        }

        if (!pageItem.pageItems) return;
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            collectTextFramesFromItem(pageItem.pageItems[i], foundFrames);
        }
    }

    /**
     * アイテム群の中のテキストの内容を集める
     * @param {PageItem[]} pageItems - 対象のアイテム
     * @returns {string[]} テキストの内容
     */
    function extractTextContentsFromItems(pageItems) {
        var textFrames = [];
        var textContents = [];

        for (var i = 0; i < pageItems.length; i++) {
            collectTextFramesFromItem(pageItems[i], textFrames);
        }

        for (var j = 0; j < textFrames.length; j++) {
            textContents.push(textFrames[j].contents);
        }

        return textContents;
    }

    /**
     * アイテムを後ろから順に削除する
     * @param {PageItem[]} pageItems - 削除するアイテム
     * @returns {void}
     */
    function removeItems(pageItems) {
        for (var i = pageItems.length - 1; i >= 0; i--) {
            /* すでに消えたアイテムは飛ばす / skip items that are already gone */
            try {
                if (pageItems[i] && pageItems[i].isValid !== false) {
                    pageItems[i].remove();
                }
            } catch (e) { }
        }
    }

    /**
     * シンボルインスタンスを一時レイヤーに複製してリンクを解除し、中のテキストを読む
     * @param {SymbolItem} symbolItem - 対象のシンボルインスタンス
     * @param {Layer} tempLayer - 一時レイヤー
     * @returns {string[]} 中のテキストの内容（読めなければ空）
     */
    function readSymbolItemTexts(symbolItem, tempLayer) {
        var symbolTexts = [];
        /* 複製・リンク解除に失敗したシンボルは読み飛ばす / skip symbols that cannot be duplicated or expanded */
        try {
            var workingSymbolItem = symbolItem.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING);
            workingSymbolItem.selected = false;
            clearDocumentSelection();
            workingSymbolItem.selected = true;
            workingSymbolItem.breakLink();

            /* リンク解除で生まれたアイテムは選択されている / breakLink leaves the expanded items selected */
            var brokenItems = getSelectedItems();
            if (brokenItems.length > 0) {
                symbolTexts = extractTextContentsFromItems(brokenItems);
                removeItems(brokenItems);
            }
        } catch (e) {
        } finally {
            clearDocumentSelection();
        }
        return symbolTexts;
    }

    /**
     * 一覧用に、シンボルの中のテキストを集める（同じシンボルは1回だけ）
     * @param {string} collectMode - "all" / "allArtboards" / "current"
     * @returns {string[]} 改行をスペースにしたテキスト
     */
    function collectSymbolTexts(collectMode) {
        var symbolTexts = [];
        var symbolItems = doc.symbolItems;
        var symbolItemsToRead = [];
        var pickedSymbolNames = {};

        for (var i = 0; i < symbolItems.length; i++) {
            var symbolItem = symbolItems[i];
            if (!symbolItem || symbolItem.isValid === false) continue;
            var symbolName = getSymbolName(symbolItem);
            if (!symbolName || pickedSymbolNames[symbolName]) continue;
            if (!isInArtboardScope(getItemArtboardIndex(symbolItem), collectMode)) continue;

            pickedSymbolNames[symbolName] = true;
            symbolItemsToRead.push(symbolItem);
        }

        if (symbolItemsToRead.length === 0) {
            return symbolTexts;
        }

        withTempSymbolLayer(function (tempLayer) {
            for (var j = symbolItemsToRead.length - 1; j >= 0; j--) {
                var readItem = symbolItemsToRead[j];
                if (!readItem || readItem.isValid === false) continue;
                var itemTexts = readSymbolItemTexts(readItem, tempLayer);
                for (var k = 0; k < itemTexts.length; k++) {
                    symbolTexts.push(itemTexts[k].replace(/[\r\n]+/g, ' '));
                }
            }
        });

        return symbolTexts;
    }

    /**
     * シンボルインスタンスがレイヤーの設定・非表示・ロックの条件で読んでよいか判定する
     * @param {SymbolItem} symbolItem - 対象のシンボルインスタンス
     * @param {Object} collectOptions - 収集の設定
     * @returns {boolean} 読んでよければ true
     */
    function isReadableSymbolItem(symbolItem, collectOptions) {
        if (!symbolItem || symbolItem.isValid === false) return false;
        if (!collectOptions.includeComment && isCommentLayer(symbolItem)) return false;
        return isCollectable(symbolItem, collectOptions);
    }

    /**
     * 書き出し用に、シンボルの中のテキストをアートボード番号付きで集める（同じシンボルは1回だけ）
     * @param {Object} collectOptions - mode / includeComment / includeLocked / includeHidden を持つ設定
     * @returns {{artboardIndex: number, text: string}[]} テキストとアートボード番号
     */
    function collectSymbolTextsByArtboard(collectOptions) {
        var symbolEntries = [];
        var symbolItems = doc.symbolItems;
        /* true = これから読む、false = 読んだ（または読まない） / true: to read, false: done or skipped */
        var symbolNameStates = {};

        for (var i = 0; i < symbolItems.length; i++) {
            var symbolItem = symbolItems[i];
            if (!isReadableSymbolItem(symbolItem, collectOptions)) continue;
            var symbolName = getSymbolName(symbolItem);
            if (!symbolName || symbolNameStates[symbolName]) continue;
            if (!isInArtboardScope(getItemArtboardIndex(symbolItem), collectOptions.mode)) continue;

            symbolNameStates[symbolName] = true;
        }

        withTempSymbolLayer(function (tempLayer) {
            for (var j = 0; j < symbolItems.length; j++) {
                var sourceItem = symbolItems[j];
                if (!isReadableSymbolItem(sourceItem, collectOptions)) continue;
                var sourceSymbolName = getSymbolName(sourceItem);
                if (!sourceSymbolName || symbolNameStates[sourceSymbolName] !== true) continue;

                /* 同じシンボルは最初に出会ったインスタンスだけを見る。それが範囲外なら、そのシンボルは読まない
                   Only the first instance met is considered; if it is out of scope, the symbol is skipped */
                symbolNameStates[sourceSymbolName] = false;
                var sourceArtboardIndex = getItemArtboardIndex(sourceItem);
                if (!isInArtboardScope(sourceArtboardIndex, collectOptions.mode)) continue;

                var itemTexts = readSymbolItemTexts(sourceItem, tempLayer);
                for (var k = 0; k < itemTexts.length; k++) {
                    symbolEntries.push({
                        artboardIndex: sourceArtboardIndex,
                        text: itemTexts[k]
                    });
                }
            }
        });

        return symbolEntries;
    }

    // =========================================
    // 一覧と書き出しのテキスト / List and export text
    // =========================================

    /* 一覧用のシンボルテキストのキャッシュ（シンボルの展開は重い） / Cache of symbol texts; expanding symbols is slow */
    var symbolTextCache = null;
    var symbolTextCacheKey = null;

    /**
     * 一覧用のシンボルテキストを返す（設定が同じならキャッシュを使う）
     * @param {Object} collectOptions - 収集の設定
     * @returns {string[]} シンボルの中のテキスト
     */
    function getSymbolTexts(collectOptions) {
        var cacheKey = [
            collectOptions.mode,
            collectOptions.includeComment ? '1' : '0',
            collectOptions.includeLocked ? '1' : '0',
            collectOptions.includeHidden ? '1' : '0'
        ].join('|');

        if (symbolTextCache !== null && symbolTextCacheKey === cacheKey) {
            return symbolTextCache;
        }

        symbolTextCache = collectSymbolTexts(collectOptions.mode);
        symbolTextCacheKey = cacheKey;
        return symbolTextCache;
    }

    /**
     * アートボードの見出し（「アートボード1: 名前」）を返す
     * @param {number} artboardIndex - アートボードの番号
     * @returns {string} 見出しの文字列
     */
    function getArtboardDisplayName(artboardIndex) {
        var artboardName = '';
        /* 名前を読めないアートボードは番号で代える / fall back to the number when the name cannot be read */
        try {
            artboardName = doc.artboards[artboardIndex].name;
        } catch (e) { }
        if (!artboardName) {
            artboardName = (uiLang === 'ja') ? 'アートボード ' + (artboardIndex + 1) : 'Artboard ' + (artboardIndex + 1);
        }
        var numberLabel = (uiLang === 'ja') ? ('アートボード' + (artboardIndex + 1)) : ('Artboard ' + (artboardIndex + 1));
        return numberLabel + ': ' + artboardName;
    }

    /**
     * 書き出すテキストをアートボードごとにまとめる
     * @param {Object} collectOptions - mode / includeComment / includeLocked / includeHidden を持つ設定
     * @returns {{name: string, texts: string[]}[]} アートボードごとの見出しとテキスト（アートボード外は最後）
     */
    function collectArtboardGroupedExportData(collectOptions) {
        var artboardGroups = [];
        var outsideGroup = {
            name: (uiLang === 'ja') ? 'アートボード外' : 'Outside Artboards',
            texts: []
        };
        var textFrames = collectDocumentTextFrames(collectOptions);
        var symbolEntries = collectSymbolTextsByArtboard(collectOptions);
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        var i;

        if (collectOptions.mode === 'current') {
            artboardGroups.push({ name: getArtboardDisplayName(activeIndex), texts: [] });
        } else {
            for (i = 0; i < doc.artboards.length; i++) {
                artboardGroups.push({ name: getArtboardDisplayName(i), texts: [] });
            }
        }

        /**
         * アートボード番号に合った見出しの下へテキストを入れる
         * @param {number} artboardIndex - アートボードの番号（-1 はアートボード外）
         * @param {string} exportText - 入れるテキスト
         * @returns {void}
         */
        function addExportText(artboardIndex, exportText) {
            if (collectOptions.mode === 'current') {
                if (artboardIndex === activeIndex) {
                    artboardGroups[0].texts.push(exportText);
                }
            } else if (artboardIndex >= 0) {
                artboardGroups[artboardIndex].texts.push(exportText);
            } else {
                outsideGroup.texts.push(exportText);
            }
        }

        for (i = 0; i < textFrames.length; i++) {
            addExportText(getItemArtboardIndex(textFrames[i]), textFrames[i].contents);
        }
        for (i = 0; i < symbolEntries.length; i++) {
            addExportText(symbolEntries[i].artboardIndex, symbolEntries[i].text);
        }

        if (collectOptions.mode !== 'current' && outsideGroup.texts.length > 0) {
            artboardGroups.push(outsideGroup);
        }
        return artboardGroups;
    }

    /**
     * 書き出すテキスト（アートボードごとに ---見出し--- を付けたもの）を作る
     * @param {Object} collectOptions - 収集の設定
     * @returns {string} ファイルに書くテキスト
     */
    function buildExportText(collectOptions) {
        var artboardGroups = collectArtboardGroupedExportData(collectOptions);
        var exportLines = [];

        for (var i = 0; i < artboardGroups.length; i++) {
            if (i > 0) exportLines.push('');
            exportLines.push('---' + artboardGroups[i].name + '---');
            for (var j = 0; j < artboardGroups[i].texts.length; j++) {
                exportLines.push(artboardGroups[i].texts[j]);
            }
        }

        return exportLines.join('\n');
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 縦並びのパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} titlePath - 見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentGroup, titlePath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * ラジオボタンかチェックボックスを、LABELS のキーで tooltip 付きで追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelKey - LABELS.radio（または LABELS.checkbox）と LABELS.tooltip のキー
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? "radio." : "checkbox.";
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelCategory + labelKey));
        addedControl.helpTip = getLabel("tooltip." + labelKey);
        return addedControl;
    }

    /**
     * ダイアログを組み立てる（イベントの配線は main() で行う）
     * @returns {Object} 作成したダイアログとコントロール
     */
    function buildDialogUI() {
        var exportDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        exportDialog.orientation = "column";
        exportDialog.alignChildren = ["fill", "top"];

        /* 2カラム: 左=テキスト一覧、右=オプション、下=ボタンエリア / Two columns: left = text list, right = options, bottom = button area */
        var columnsGroup = exportDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "fill"];

        /* 左カラム / Left column */
        var listColumn = columnsGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];

        listColumn.add("statictext", undefined, getLabel("fieldLabel.textList"));
        var textListField = listColumn.add("edittext", TEXT_LIST_BOUNDS, "", { multiline: true, scrolling: true, readonly: true });
        textListField.helpTip = getLabel("tooltip.textList");

        /* 右カラム / Right column */
        var optionsColumn = columnsGroup.add("group");
        optionsColumn.orientation = "column";
        optionsColumn.alignChildren = ["fill", "top"];

        /* 対象テキストパネル（アートボード） / Target text panel (artboards) */
        var artboardPanel = addOptionPanel(optionsColumn, "panel.targetArtboards");

        var scopeRadioGroup = artboardPanel.add("group");
        scopeRadioGroup.orientation = "column";
        scopeRadioGroup.alignChildren = ["left", "top"];
        var rbCurrentArtboard = addLabeledControl(scopeRadioGroup, "radiobutton", "currentArtboard");
        var rbAllArtboards = addLabeledControl(scopeRadioGroup, "radiobutton", "allArtboards");
        rbCurrentArtboard.value = true;

        var outsideRow = artboardPanel.add("group");
        outsideRow.orientation = "row";
        var cbIncludeOutside = addLabeledControl(outsideRow, "checkbox", "includeOutside");
        cbIncludeOutside.value = false;
        cbIncludeOutside.enabled = false;

        /* 対象テキストパネル（レイヤー） / Target text panel (layers) */
        var layerPanel = addOptionPanel(optionsColumn, "panel.targetLayers");
        var cbIncludeComment = addLabeledControl(layerPanel, "checkbox", "includeCommentLayers");
        cbIncludeComment.value = false;
        var cbIncludeLocked = addLabeledControl(layerPanel, "checkbox", "includeLocked");
        cbIncludeLocked.value = false;
        var cbIncludeHidden = addLabeledControl(layerPanel, "checkbox", "includeHidden");
        cbIncludeHidden.value = false;

        var duplicateRow = optionsColumn.add("group");
        duplicateRow.orientation = "row";
        duplicateRow.alignChildren = ["center", "center"];
        duplicateRow.alignment = ["center", "top"];
        var cbRemoveDuplicates = addLabeledControl(duplicateRow, "checkbox", "removeDuplicates");
        cbRemoveDuplicates.value = false;

        /* ボタンエリア / Button area */
        var btnRowGroup = exportDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnCancel = btnLeftGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCopyText = btnRightGroup.add("button", undefined, getLabel("button.copyText"));
        btnCopyText.helpTip = getLabel("tooltip.copyText");
        var btnExportText = btnRightGroup.add("button", undefined, getLabel("button.exportText"), { name: "ok" });
        btnExportText.helpTip = getLabel("tooltip.exportText");
        exportDialog.defaultElement = btnExportText;

        return {
            exportDialog: exportDialog,
            rbCurrentArtboard: rbCurrentArtboard,
            rbAllArtboards: rbAllArtboards,
            cbIncludeOutside: cbIncludeOutside,
            cbRemoveDuplicates: cbRemoveDuplicates,
            cbIncludeComment: cbIncludeComment,
            cbIncludeLocked: cbIncludeLocked,
            cbIncludeHidden: cbIncludeHidden,
            textListField: textListField,
            btnCancel: btnCancel,
            btnCopyText: btnCopyText,
            btnExportText: btnExportText
        };
    }

    /**
     * ダイアログの設定を読む
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @returns {{mode: string, includeComment: boolean, includeLocked: boolean, includeHidden: boolean}} 収集の設定
     */
    function readCollectOptions(dialogControls) {
        var collectMode;
        if (dialogControls.cbIncludeOutside.value) {
            collectMode = "all";
        } else if (dialogControls.rbAllArtboards.value) {
            collectMode = "allArtboards";
        } else {
            collectMode = "current";
        }
        return {
            mode: collectMode,
            includeComment: dialogControls.cbIncludeComment.value,
            includeLocked: dialogControls.cbIncludeLocked.value,
            includeHidden: dialogControls.cbIncludeHidden.value
        };
    }

    /**
     * 設定に合わせてテキストを集め直し、一覧を更新する
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateTextList(dialogControls) {
        var collectOptions = readCollectOptions(dialogControls);
        var listEntries = collectDocumentTextFrames(collectOptions);
        var symbolTexts = getSymbolTexts(collectOptions);
        for (var i = 0; i < symbolTexts.length; i++) {
            listEntries.push({ contents: symbolTexts[i] });
        }
        if (dialogControls.cbRemoveDuplicates.value) {
            listEntries = removeDuplicateTextEntries(listEntries);
        }
        var listLines = [];
        for (var j = 0; j < listEntries.length; j++) {
            listLines.push(makeListLine(listEntries[j].contents));
        }
        dialogControls.textListField.text = listLines.join("\n");
    }

    /**
     * テキストをデスクトップのファイルへ書き出す
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @returns {void}
     */
    function exportTextToDesktop(dialogControls) {
        var filePath;
        var exportFile = null;
        var exportText;
        /* ファイルの I/O / file I/O */
        try {
            exportText = buildExportText(readCollectOptions(dialogControls));
            filePath = Folder.desktop.fsName + '/text-' + sanitizeFileName(getDocumentBaseName(doc)) + '-' + getDateStamp() + '.txt';
            exportFile = new File(filePath);
            exportFile.encoding = 'UTF-8';
            exportFile.lineFeed = 'Unix';
            if (!exportFile.open('w')) {
                throw new Error('open failed: ' + filePath);
            }
            exportFile.write(exportText);
            exportFile.close();
            alert(getLabel('alert.exportDone') + '\n' + exportFile.fsName);
            dialogControls.exportDialog.close();
        } catch (e) {
            try {
                if (exportFile && exportFile.opened) exportFile.close();
            } catch (closeErr) { }
            alert(getLabel('alert.exportFailed') + '\n' + e);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示して、テキストの一覧・コピー・書き出しを行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        doc = app.activeDocument;

        var dialogControls = buildDialogUI();
        var textListField = dialogControls.textListField;
        var cbIncludeOutside = dialogControls.cbIncludeOutside;

        /**
         * 一覧を更新する（イベントハンドラ用）
         * @returns {void}
         */
        function refreshTextList() {
            updateTextList(dialogControls);
        }

        dialogControls.btnCopyText.onClick = function () {
            /* メニューコマンドのコピー / menu command copy */
            try {
                withSelectionSaved(function () {
                    textListField.active = true;
                    textListField.textselection = textListField.text;
                    app.executeMenuCommand("copy");
                }, true);
            } catch (e) {
                alert(e);
            }
        };

        dialogControls.btnExportText.onClick = function () {
            exportTextToDesktop(dialogControls);
        };

        /* 初回収集 / Initial collection */
        refreshTextList();

        /* ラジオボタン・チェックボックス切り替え時に更新 / Refresh when radio buttons or checkboxes change */
        dialogControls.rbCurrentArtboard.onClick = function () {
            cbIncludeOutside.enabled = false;
            cbIncludeOutside.value = false;
            refreshTextList();
        };
        dialogControls.rbAllArtboards.onClick = function () {
            cbIncludeOutside.enabled = true;
            refreshTextList();
        };

        cbIncludeOutside.onClick = refreshTextList;
        dialogControls.cbRemoveDuplicates.onClick = refreshTextList;
        dialogControls.cbIncludeComment.onClick = refreshTextList;
        dialogControls.cbIncludeLocked.onClick = refreshTextList;
        dialogControls.cbIncludeHidden.onClick = refreshTextList;

        /* キャンセルボタンで閉じる / Close the dialog when Cancel is pressed */
        dialogControls.btnCancel.onClick = function () {
            dialogControls.exportDialog.close();
        };

        dialogControls.exportDialog.show();
    }

    main();

})();
