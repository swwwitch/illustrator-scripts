#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストを収集して一覧表示し、その場で編集してドキュメントへ書き戻します。
対象はアートボード単位やレイヤー単位で絞り込めます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextScopeEdit.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb845889dd553

### Overview

Collects the text in the document, lists it, and lets you edit it in place and write it back.
The scope can be narrowed by artboard or by layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextScopeEdit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextScopeEdit";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextScopeEdit.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextScopeEdit.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb845889dd553"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var SOFT_BREAK              = '@#';  /* 編集欄での強制改行の代替文字 / Placeholder for forced line breaks in the edit field */
    var POSITION_SORT_TOLERANCE = 10;    /* 位置順で同じ行とみなす高さの差（pt）/ Height difference treated as the same row (pt) */
    var LIST_LABEL_MAX_LENGTH   = 40;    /* 一覧に出す文字数 / Characters shown per list row */

    // =========================================
    // 一時レイヤー / Temporary layer
    // =========================================

    /* シンボル内テキストを読むための作業レイヤー（名前とメモで見分ける）
       Work layer for reading text in symbols (identified by name and note) */
    var TEMP_LAYER_NAME = '__TextScopeEdit_temp_read__';
    var TEMP_LAYER_NOTE = '__TextScopeEdit_temp_read__';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var TABS_SIZE                = [550, 350];         /* タブ全体の大きさ / Tabbed panel size */
    var TABS_MARGINS             = [15, 10, 1, 10];    /* タブ全体の余白 [左,上,右,下] / Tabbed panel margins */
    var CANVAS_TAB_MARGINS       = [15, 20, 1, 10];    /* ［カンバス］タブの余白 / Canvas tab margins */
    var CANVAS_TAB_SPACING       = 15;                 /* ［カンバス］タブの要素間隔 / Canvas tab spacing */
    var INFO_TAB_MARGINS         = [15, 20, 0, 10];    /* 名前一覧タブの余白 / Name-list tab margins */
    var INFO_TAB_SPACING         = 10;                 /* 名前一覧タブの要素間隔 / Name-list tab spacing */
    var CHOICE_ROW_SPACING       = 15;                 /* ラジオ・チェックボックスの横並びの間隔 / Spacing of radio/checkbox rows */
    var SCOPE_PANEL_MARGINS      = [15, 20, 1, 10];    /* 右カラムのパネル余白 / Right-column panel margins */
    var OPTIONS_GROUP_MARGINS    = [15, 5, 15, 10];    /* オプション群の余白 / Option group margins */
    var EXPORT_PANEL_MARGINS     = [15, 20, 15, 10];   /* 書き出しオプションのパネル余白 / Export options panel margins */
    var TEXT_LIST_BOUNDS         = [0, 0, 250, 194];   /* テキスト一覧の大きさ / Text list bounds */
    var TEXT_EDIT_BOUNDS         = [0, 0, 250, 72];    /* テキスト編集欄の大きさ / Text edit field bounds */
    var SYMBOL_LIST_WIDTH        = 250;                /* シンボル内テキスト一覧の幅 / Symbol text list width */
    var SYMBOL_LIST_MIN_ROWS     = 4;                  /* シンボル内テキスト一覧の最小行数 / Minimum rows */
    var SYMBOL_LIST_MAX_ROWS     = 8;                  /* シンボル内テキスト一覧の最大行数 / Maximum rows */
    var SYMBOL_LIST_ROW_HEIGHT   = 18;                 /* 1行の高さ / Row height */
    var SYMBOL_LIST_EXTRA_HEIGHT = 6;                  /* 行以外に足す高さ / Extra height */
    var FONT_COLUMN_WIDTHS       = [180, 180, 120];    /* フォント一覧の列幅 / Font list column widths */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "テキストの収集と編集", en: "Collect and Edit Text" },
            exportOptions: { ja: "書き出しオプション", en: "Export Options" }
        },
        tab: {
            canvas: { ja: "カンバス", en: "Canvas" },
            layerNames: { ja: "レイヤー名", en: "Layer Names" },
            artboardNames: { ja: "アートボード名", en: "Artboard Names" },
            fontNames: { ja: "フォント名", en: "Font Name" }
        },
        panel: {
            targetText: { ja: "対象テキスト（アートボード）", en: "Text Scope (Artboards)" },
            layerText: { ja: "対象テキスト（レイヤーなど）", en: "Text Scope (Layers)" },
            sort: { ja: "ソート", en: "Sort" }
        },
        radio: {
            currentArtboard: { ja: "現在のアートボード内", en: "Current Artboard Only" },
            allArtboards: { ja: "すべてのアートボード内", en: "All Artboards Only" },
            sortNone: { ja: "なし", en: "None" },
            sortPosition: { ja: "位置順", en: "Sort by Position" },
            sortAlphabetical: { ja: "ABC順", en: "Sort Alphabetically" },
            layerScopeAll: { ja: "すべてのレイヤー", en: "All Layers" },
            layerScopeTop: { ja: "上位レベルのレイヤーのみ", en: "Top-Level Layers Only" },
            artboardScopeNumbered: { ja: "番号つき", en: "Numbered" },
            artboardScopeRaw: { ja: "アートボード名のみ", en: "Names Only" }
        },
        checkbox: {
            mergeDuplicates: { ja: "同じ内容を一括編集", en: "Treat Same Text as One" },
            wholeDocument: { ja: "ドキュメント全体を対象に", en: "Include Entire Document" },
            includeCommentLayers: { ja: "//ではじめるレイヤーも含む", en: "Include Layers That Start with //" },
            includeLocked: { ja: "ロックされたテキスト", en: "Locked Text" },
            includeHidden: { ja: "非表示のテキスト", en: "Hidden Text" },
            keepFormat: { ja: "段落書式を保持", en: "Keep Paragraph Formatting" },
            preview: { ja: "プレビュー", en: "Preview" },
            exportIncludeText: { ja: "テキスト", en: "Text" },
            exportIncludeFonts: { ja: "フォント名", en: "Font Names" },
            exportOpenAfter: { ja: "書き出し後にファイルを開く", en: "Open file after export" },
            fontPS: { ja: "PostScript名", en: "PostScript Name" },
            fontFamily: { ja: "フォント名", en: "Font Name" },
            fontStyle: { ja: "スタイル", en: "Style" }
        },
        fieldLabel: {
            textList: { ja: "テキスト一覧", en: "Text List" },
            textEdit: { ja: "テキスト編集", en: "Edit Text" },
            symbolText: { ja: "シンボル内テキスト（編集不可）", en: "Text in Symbols (Read-Only)" },
            layerNameList: { ja: "レイヤー名一覧", en: "Layer Name List" },
            artboardNameList: { ja: "アートボード名一覧", en: "Artboard Name List" },
            fontList: { ja: "フォント一覧", en: "Font List" }
        },
        listColumn: {
            fontPS: { ja: "PostScript名", en: "PostScript Name" },
            fontFamily: { ja: "フォント名", en: "Font Name" },
            fontStyle: { ja: "スタイル", en: "Style" }
        },
        format: {
            itemPrefix: { ja: ": ", en: ": " },
            symbolSeparator: { ja: "：", en: ": " },
            symbolSuffix: { ja: "〈シンボル：{symbolName}〉", en: " «Symbol: {symbolName}»" }
        },
        tooltip: {
            keepFormat: { ja: "段落の書式を保持したまま\nテキストを置換します", en: "Replace text while preserving\nparagraph formatting" },
            preview: { ja: "編集結果をリアルタイムで\nプレビューします", en: "Preview the edited result\nin real time" },
            previewDisabled: { ja: "CC 2020ではプレビュー無効", en: "Preview is disabled in CC 2020" },
            wholeDocument: {
                ja: "アートボードの外にあるテキストも対象にします\n（「すべてのアートボード内」のときに選べます）",
                en: "Also includes text outside the artboards\n(available with All Artboards Only)"
            },
            mergeDuplicates: {
                ja: "同じ内容のテキストを一覧の1行にまとめ、\n編集をそのすべてに反映します",
                en: "Lists identical text as one row and applies\nthe edit to every copy"
            },
            sortPosition: {
                ja: "上から下へ並べ、ほぼ同じ高さのものは\n左から右へ並べます",
                en: "Sorts top to bottom, and left to right\nfor text at about the same height"
            },
            sortAlphabetical: { ja: "大文字と小文字を区別せずに\n内容の順に並べます", en: "Sorts by content, ignoring case" },
            textEdit: {
                ja: "Shift+Enter で強制改行を入れます\n（{softBreak} と表示されます）",
                en: "Shift+Enter inserts a forced line break\n(shown as {softBreak})"
            },
            exportText: {
                ja: "アートボードごとにテキストとフォント名をまとめ、\nデスクトップにテキストファイルで書き出します",
                en: "Writes the text and font names, grouped by artboard,\nto a text file on the desktop"
            },
            fontList: {
                ja: "行をクリックすると、そのフォントを使っている\nテキストを選択します",
                en: "Click a row to select the text\nthat uses that font"
            },
            includeLocked: {
                ja: "ロックされたテキストと、ロックされた\nレイヤー・グループ内のテキストも対象にします",
                en: "Also includes locked text and text\nin locked layers or groups"
            },
            includeHidden: {
                ja: "非表示のテキストと、非表示の\nレイヤー・グループ内のテキストも対象にします",
                en: "Also includes hidden text and text\nin hidden layers or groups"
            }
        },
        button: {
            exportText: { ja: "テキスト書き出し", en: "Export Text" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        fallbackName: {
            artboardNumber: { ja: "アートボード{number}", en: "Artboard {number}" },
            outsideArtboards: { ja: "アートボード外", en: "Outside Artboards" },
            unknownFont: { ja: "不明", en: "Unknown" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
            exportFailed: { ja: "テキストを書き出せませんでした", en: "Failed to export text" },
            noMatchingFontText: { ja: "該当するテキストがありません", en: "No matching text found" }
        }
    };

    /**
     * ドット区切りのパスで LABELS から表示言語の文字列を引く
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        return labelNode[uiLang];
    }

    // =========================================
    // 文字列の整形 / String helpers
    // =========================================

    /**
     * 一覧用のラベルを作る（改行を空白にし、長ければ先頭だけにする）
     * @param {string} text - テキストの内容
     * @param {number} [maxLength] - 表示する最大文字数（省略時は LIST_LABEL_MAX_LENGTH）
     * @returns {string} 一覧用のラベル
     */
    function buildListLabel(text, maxLength) {
        if (!maxLength) maxLength = LIST_LABEL_MAX_LENGTH;
        var labelText = text.replace(/[\r\n]+/g, " ");
        if (labelText.length > maxLength) labelText = labelText.substring(0, maxLength) + "…";
        return labelText;
    }

    /**
     * 2桁にゼロ埋めする
     * @param {number} num - 数値
     * @returns {string} 2桁の文字列
     */
    function zeroPad2(num) {
        return (num < 10 ? '0' : '') + num;
    }

    /**
     * 現在日時を YYYYMMDD-HHMMSS 形式で返す
     * @returns {string} 日時の文字列
     */
    function getDateTimeStamp() {
        var now = new Date();
        return now.getFullYear()
            + zeroPad2(now.getMonth() + 1)
            + zeroPad2(now.getDate())
            + '-'
            + zeroPad2(now.getHours())
            + zeroPad2(now.getMinutes())
            + zeroPad2(now.getSeconds());
    }

    /**
     * ドキュメント名から拡張子を除いた名前を返す
     * @param {Document} documentRef - 対象のドキュメント
     * @returns {string} 拡張子なしの名前（無ければ "untitled"）
     */
    function getDocumentBaseName(documentRef) {
        var documentName = documentRef && documentRef.name ? documentRef.name : 'untitled';
        return documentName.replace(/\.[^\.]+$/, '');
    }

    /**
     * ファイル名に使えない文字を _ に置き換える
     * @param {string} fileName - ファイル名
     * @returns {string} 置き換えたファイル名
     */
    function sanitizeFileName(fileName) {
        return fileName.replace(/[\\\/\:\*\?\"\<\>\|]+/g, '_');
    }

    // =========================================
    // 収集の条件 / Collection filters
    // =========================================

    /**
     * 実質的に空のテキストか（改行・空白だけのものも空とみなす）
     * @param {TextFrame} textFrame - 調べるテキストフレーム
     * @returns {boolean} 空なら true
     */
    function isEmptyTextFrame(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== "TextFrame") return true;
            var frameText = textFrame.contents;
            if (!frameText) return true;
            frameText = frameText.replace(/[\r\n\x03]/g, "").replace(/\s+/g, "");
            return frameText.length === 0;
        } catch (e) {
            return true;
        }
    }

    /**
     * レイヤー名が // ではじまるか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} // ではじまるレイヤーにあれば true
     */
    function isCommentLayer(pageItem) {
        try {
            return pageItem.layer.name.indexOf("//") === 0;
        } catch (err) {
            return false;
        }
    }

    /**
     * 親のグループ・レイヤーのどれかがロックまたは非表示か
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @returns {boolean} 除外すべき親があれば true
     */
    function isParentLockedOrHidden(pageItem, collectOptions) {
        var parentItem;
        try {
            parentItem = pageItem ? pageItem.parent : null;
        } catch (e) {
            parentItem = null;
        }

        while (parentItem) {
            try {
                if (parentItem.typename === 'Document') break;
                if (!collectOptions.includeLocked && parentItem.locked) return true;
                if (!collectOptions.includeHidden) {
                    if (parentItem.hidden) return true;
                    if (parentItem.visible === false) return true;
                }
            } catch (e2) { }

            try {
                parentItem = parentItem.parent;
            } catch (e3) {
                parentItem = null;
            }
        }
        return false;
    }

    /**
     * 非表示・ロックの設定に照らして収集してよいか（親のグループ・レイヤーも見る）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @returns {boolean} 収集してよければ true
     */
    function isCollectable(pageItem, collectOptions) {
        try {
            if (!pageItem) return false;
            if (!collectOptions.includeHidden && pageItem.hidden) return false;
            if (!collectOptions.includeLocked && pageItem.locked) return false;
            if (pageItem.layer) {
                if (!collectOptions.includeLocked && pageItem.layer.locked) return false;
                if (!collectOptions.includeHidden && pageItem.layer.visible === false) return false;
                if (!collectOptions.includeHidden && pageItem.layer.hidden) return false;
            }
            if (isParentLockedOrHidden(pageItem, collectOptions)) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 直下の子かどうか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {Object} parentContainer - 親のレイヤーまたはグループ
     * @returns {boolean} 直下の子なら true
     */
    function isDirectChildOf(pageItem, parentContainer) {
        try {
            return pageItem && parentContainer && pageItem.parent === parentContainer;
        } catch (e) {
            return false;
        }
    }

    /**
     * アートボードと一部でも重なっているか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {number[]} artboardRect - アートボードの [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function isOnArtboard(pageItem, artboardRect) {
        var bounds = pageItem.geometricBounds;
        return (bounds[2] > artboardRect[0] && bounds[0] < artboardRect[2] &&
            bounds[1] > artboardRect[3] && bounds[3] < artboardRect[1]);
    }

    /**
     * いずれかのアートボードに重なっているか
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 重なっていれば true
     */
    function isOnAnyArtboard(doc, pageItem) {
        return getItemArtboardIndex(doc, pageItem) >= 0;
    }

    /**
     * 重なっている最初のアートボードの番号を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {number} アートボードの番号（0 始まり）。どこにも重ならなければ -1
     */
    function getItemArtboardIndex(doc, pageItem) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isOnArtboard(pageItem, doc.artboards[i].artboardRect)) {
                return i;
            }
        }
        return -1;
    }

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    /**
     * レイヤーまたはグループの直下からテキストを集め、グループは中へたどる
     * @param {Object} container - レイヤーまたはグループ
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームを入れる配列
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {Function} acceptFrame - テキストフレームを受け入れるか判定する関数
     * @returns {void}
     */
    function collectTextFromContainer(container, collectedFrames, collectOptions, acceptFrame) {
        if (!container || !container.pageItems) return;
        var pageItems = container.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!isDirectChildOf(pageItem, container)) continue;
            if (!collectOptions.includeComment && isCommentLayer(pageItem)) continue;
            if (!isCollectable(pageItem, collectOptions)) continue;

            if (pageItem.typename === "TextFrame") {
                if (isEmptyTextFrame(pageItem)) continue;
                if (acceptFrame(pageItem)) {
                    collectedFrames.push(pageItem);
                }
            } else if (pageItem.typename === "GroupItem") {
                collectTextFromContainer(pageItem, collectedFrames, collectOptions, acceptFrame);
            }
        }
    }

    /**
     * レイヤーとそのサブレイヤーからテキストを集める
     * @param {Layer} layer - 対象のレイヤー
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームを入れる配列
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {Function} acceptFrame - テキストフレームを受け入れるか判定する関数
     * @returns {void}
     */
    function collectTextFromLayer(layer, collectedFrames, collectOptions, acceptFrame) {
        if (!layer) return;
        collectTextFromContainer(layer, collectedFrames, collectOptions, acceptFrame);
        for (var i = 0; i < layer.layers.length; i++) {
            collectTextFromLayer(layer.layers[i], collectedFrames, collectOptions, acceptFrame);
        }
    }

    /**
     * 範囲に応じてテキストフレームを集める
     * @param {Document} doc - 対象のドキュメント
     * @param {string} scopeMode - "all"（ドキュメント全体）/ "allArtboards"（いずれかのアートボード）/ "current"（現在のアートボード）
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function collectFramesByScope(doc, scopeMode, collectOptions) {
        var collectedFrames = [];
        var acceptFrame;
        if (scopeMode === "all") {
            acceptFrame = function () { return true; };
        } else if (scopeMode === "allArtboards") {
            acceptFrame = function (textFrame) { return isOnAnyArtboard(doc, textFrame); };
        } else {
            var abIndex = doc.artboards.getActiveArtboardIndex();
            var artboardRect = doc.artboards[abIndex].artboardRect;
            acceptFrame = function (textFrame) { return isOnArtboard(textFrame, artboardRect); };
        }
        for (var i = 0; i < doc.layers.length; i++) {
            collectTextFromLayer(doc.layers[i], collectedFrames, collectOptions, acceptFrame);
        }
        return collectedFrames;
    }

    /**
     * 名前一覧タブ用に、ドキュメント全体のテキストを条件なしで集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function getScopedFramesForInfoTabs(doc) {
        return collectFramesByScope(doc, 'all', {
            includeComment: true,
            includeLocked: true,
            includeHidden: true
        });
    }

    /**
     * 同じ内容のテキストフレームをまとめる
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {{uniqueFrames: TextFrame[], duplicateMap: Array}} 内容ごとの代表と、代表ごとの同じ内容の全フレーム
     */
    function groupDuplicateFrames(textFrames) {
        var uniqueFrames = [];
        var duplicateMap = [];
        var contentIndexMap = {}; /* 内容 → uniqueFrames の番号 / contents -> index in uniqueFrames */
        for (var i = 0; i < textFrames.length; i++) {
            var frameContents = textFrames[i].contents;
            if (contentIndexMap[frameContents] === undefined) {
                contentIndexMap[frameContents] = uniqueFrames.length;
                uniqueFrames.push(textFrames[i]);
                duplicateMap.push([textFrames[i]]);
            } else {
                duplicateMap[contentIndexMap[frameContents]].push(textFrames[i]);
            }
        }
        return { uniqueFrames: uniqueFrames, duplicateMap: duplicateMap };
    }

    // =========================================
    // レイヤー名・アートボード名 / Layer and artboard names
    // =========================================

    /**
     * レイヤーとそのサブレイヤーの名前を重複なく集める
     * @param {Layer} layer - 対象のレイヤー
     * @param {string[]} layerNames - 集めた名前を入れる配列
     * @param {Object} seenNames - 既出の名前の表
     * @returns {void}
     */
    function collectAllLayerNamesFromLayer(layer, layerNames, seenNames) {
        var layerName = '';
        if (!layer) return;

        try {
            layerName = layer.name;
        } catch (e) {
            layerName = '';
        }
        if (layerName && !seenNames[layerName]) {
            seenNames[layerName] = true;
            layerNames.push(layerName);
        }

        for (var i = 0; i < layer.layers.length; i++) {
            collectAllLayerNamesFromLayer(layer.layers[i], layerNames, seenNames);
        }
    }

    /**
     * ドキュメントのすべてのレイヤー名（サブレイヤーを含む）を重複なく集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {string[]} レイヤー名の配列
     */
    function collectAllDocumentLayerNames(doc) {
        var layerNames = [];
        var seenNames = {};
        for (var i = 0; i < doc.layers.length; i++) {
            collectAllLayerNamesFromLayer(doc.layers[i], layerNames, seenNames);
        }
        return layerNames;
    }

    /**
     * 最上位のレイヤー名を集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {string[]} レイヤー名の配列
     */
    function collectTopLevelLayerNames(doc) {
        var layerNames = [];
        var layerName;
        for (var i = 0; i < doc.layers.length; i++) {
            try {
                layerName = doc.layers[i].name;
            } catch (e) {
                layerName = '';
            }
            if (layerName) {
                layerNames.push(layerName);
            }
        }
        return layerNames;
    }

    /**
     * アートボード番号の既定名（「アートボード1」「Artboard 1」）を返す
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 既定名
     */
    function getArtboardNumberName(index) {
        return getLabel('fallbackName.artboardNumber').replace('{number}', index + 1);
    }

    /**
     * アートボード名を返す（空なら既定名）
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} アートボード名
     */
    function getRawArtboardName(doc, index) {
        var artboardName = '';
        try {
            artboardName = doc.artboards[index].name;
        } catch (e) { }
        if (!artboardName) {
            artboardName = getArtboardNumberName(index);
        }
        return artboardName;
    }

    /**
     * ［アートボード名］タブの「番号つき」表示（「1: 名前」）を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 表示名
     */
    function getArtboardTabDisplayName(doc, index) {
        return (index + 1) + ': ' + getRawArtboardName(doc, index);
    }

    /**
     * 書き出し用の見出し（「アートボード1: 名前」）を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {number} index - アートボードの番号（0 始まり）
     * @returns {string} 見出し
     */
    function getArtboardDisplayName(doc, index) {
        return getArtboardNumberName(index) + ': ' + getRawArtboardName(doc, index);
    }

    // =========================================
    // フォント / Fonts
    // =========================================

    /**
     * テキストで使われているフォントの PostScript 名を重複なく集める
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string[]} PostScript 名の配列（取れなければ「不明」の1件）
     */
    function getTextFrameFontNames(textFrame) {
        var seenNames = {};
        var fontNames = [];
        var fontName;

        if (!textFrame) return fontNames;

        try {
            var frameCharacters = textFrame.characters;
            for (var i = 0; i < frameCharacters.length; i++) {
                try {
                    fontName = frameCharacters[i].characterAttributes.textFont.name;
                } catch (e) {
                    fontName = '';
                }
                if (!fontName) continue;
                if (!seenNames[fontName]) {
                    seenNames[fontName] = true;
                    fontNames.push(fontName);
                }
            }
        } catch (err) { }

        if (fontNames.length === 0) {
            fontNames.push(getLabel('fallbackName.unknownFont'));
        }
        return fontNames;
    }

    /**
     * フォントの組（PostScript 名・ファミリー名・スタイル名）の比較用キーを返す
     * @param {Object} fontTriple - { psName, familyName, styleName }
     * @returns {string} タブ区切りのキー
     */
    function getFontTripleKey(fontTriple) {
        return fontTriple.psName + '\t' + fontTriple.familyName + '\t' + fontTriple.styleName;
    }

    /**
     * テキストで使われているフォントの組（PostScript 名・ファミリー名・スタイル名）を重複なく集める
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} { psName, familyName, styleName } の配列（取れなければ「不明」の1件）
     */
    function getTextFrameFontTriples(textFrame) {
        var fontTriples = [];
        var seenKeys = {};
        var unknownFont = getLabel('fallbackName.unknownFont');

        if (!textFrame) return fontTriples;

        try {
            var frameCharacters = textFrame.characters;
            for (var i = 0; i < frameCharacters.length; i++) {
                var fontTriple;
                try {
                    var textFont = frameCharacters[i].characterAttributes.textFont;
                    fontTriple = {
                        psName: textFont.name || unknownFont,
                        familyName: textFont.family || unknownFont,
                        styleName: textFont.style || ''
                    };
                } catch (e) {
                    fontTriple = { psName: unknownFont, familyName: unknownFont, styleName: '' };
                }
                var tripleKey = getFontTripleKey(fontTriple);
                if (!seenKeys[tripleKey]) {
                    seenKeys[tripleKey] = true;
                    fontTriples.push(fontTriple);
                }
            }
        } catch (err) { }

        if (fontTriples.length === 0) {
            fontTriples.push({ psName: unknownFont, familyName: unknownFont, styleName: '' });
        }
        return fontTriples;
    }

    /**
     * 複数のテキストで使われているフォントの組を、出てきた順に重複なく集める
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {Object[]} { psName, familyName, styleName } の配列
     */
    function collectUniqueFontTriples(textFrames) {
        var seenKeys = {};
        var uniqueTriples = [];
        for (var i = 0; i < textFrames.length; i++) {
            var fontTriples = getTextFrameFontTriples(textFrames[i]);
            for (var j = 0; j < fontTriples.length; j++) {
                var tripleKey = getFontTripleKey(fontTriples[j]);
                if (!seenKeys[tripleKey]) {
                    seenKeys[tripleKey] = true;
                    uniqueTriples.push(fontTriples[j]);
                }
            }
        }
        return uniqueTriples;
    }

    /**
     * 書き出し用に、フォントの行（「PostScript名  ファミリー名  スタイル名」）を並べ替えて返す
     * @param {TextFrame[]} textFrames - テキストフレームの配列
     * @returns {string[]} フォントの行
     */
    function collectFontDisplayNamesFromFrames(textFrames) {
        var fontTriples = collectUniqueFontTriples(textFrames);
        var fontLines = [];
        for (var i = 0; i < fontTriples.length; i++) {
            fontLines.push(fontTriples[i].psName + '  ' + fontTriples[i].familyName + (fontTriples[i].styleName ? ('  ' + fontTriples[i].styleName) : ''));
        }
        fontLines.sort();
        return fontLines;
    }

    /**
     * 指定のフォントを使っているテキストを選択する（ドキュメント全体から探す）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} targetPSName - フォントの PostScript 名
     * @returns {void}
     */
    function selectTextFramesByFontMatch(doc, targetPSName) {
        var textFrames = getScopedFramesForInfoTabs(doc);
        var matchedFrames = [];
        var i;

        if (!targetPSName) return;

        for (i = 0; i < textFrames.length; i++) {
            if (getTextFrameFontNames(textFrames[i]).join('\n').indexOf(targetPSName) >= 0) {
                matchedFrames.push(textFrames[i]);
            }
        }

        try {
            doc.selection = null;
        } catch (e1) { }

        if (matchedFrames.length === 0) {
            alert(getLabel('alert.noMatchingFontText'));
            return;
        }

        for (i = 0; i < matchedFrames.length; i++) {
            try {
                matchedFrames[i].selected = true;
            } catch (e2) { }
        }

        app.redraw();
    }

    // =========================================
    // 段落書式を保持した書き換え / Rewriting while keeping paragraph formatting
    // =========================================

    /* 保持する段落のプロパティ / Paragraph properties to keep */
    var PARAGRAPH_FORMAT_KEYS = [
        'justification', 'firstLineIndent', 'leftIndent', 'rightIndent',
        'spaceBefore', 'spaceAfter', 'hyphenation', 'hyphenationZone',
        'desiredWordSpacing', 'minimumWordSpacing', 'maximumWordSpacing',
        'desiredLetterSpacing', 'minimumLetterSpacing', 'maximumLetterSpacing',
        'desiredGlyphScaling', 'minimumGlyphScaling', 'maximumGlyphScaling',
        'singleWordJustification', 'everyLineComposer', 'kinsokuOrder',
        'bunriKinshi', 'kurikaeshiMojiShori', 'romanHanging', 'mojikumi',
        'kinsoku', 'leadingType'
    ];

    /**
     * オブジェクトのプロパティを読める分だけ控える
     * @param {Object} sourceObject - 読み取り元
     * @param {string[]} propertyNames - プロパティ名
     * @returns {Object} 控えたプロパティ
     */
    function readProperties(sourceObject, propertyNames) {
        var propertyValues = {};
        for (var i = 0; i < propertyNames.length; i++) {
            var propertyName = propertyNames[i];
            try {
                propertyValues[propertyName] = sourceObject[propertyName];
            } catch (err) { }
        }
        return propertyValues;
    }

    /**
     * 控えたプロパティを書ける分だけ書き戻す（undefined は飛ばす）
     * @param {Object} propertyValues - readProperties() の結果
     * @param {Object} targetObject - 書き込み先
     * @param {string[]} propertyNames - プロパティ名
     * @returns {void}
     */
    function writeProperties(propertyValues, targetObject, propertyNames) {
        for (var i = 0; i < propertyNames.length; i++) {
            var propertyName = propertyNames[i];
            if (propertyValues[propertyName] === undefined) continue;
            try {
                targetObject[propertyName] = propertyValues[propertyName];
            } catch (err) { }
        }
    }

    /**
     * テキストの段落を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object} 段落のコレクション（ストーリーが無ければ空配列）
     */
    function getParagraphs(textFrame) {
        if (!textFrame || !textFrame.story) return [];
        return textFrame.paragraphs;
    }

    /**
     * オブジェクトとレイヤーを一時的に編集できる状態にして処理し、終わったら元に戻す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Function} editAction - 編集できる状態で行う処理
     * @returns {void}
     */
    function withTemporarilyEditableItem(pageItem, editAction) {
        var targetLayer = null;
        var originalItemLocked = null;
        var originalItemHidden = null;
        var originalLayerLocked = null;
        var originalLayerVisible = null;

        try {
            if (!pageItem) return;
            targetLayer = pageItem.layer ? pageItem.layer : null;

            originalItemLocked = pageItem.locked;
            originalItemHidden = pageItem.hidden;
            if (targetLayer) {
                originalLayerLocked = targetLayer.locked;
                originalLayerVisible = targetLayer.visible;
            }

            if (targetLayer) {
                if (targetLayer.visible === false) targetLayer.visible = true;
                if (targetLayer.locked) targetLayer.locked = false;
            }
            if (pageItem.hidden) pageItem.hidden = false;
            if (pageItem.locked) pageItem.locked = false;

            editAction();
        } finally {
            /* 1つ戻せなくても残りは戻す / Keep restoring the rest even if one fails */
            try {
                if (pageItem && originalItemLocked !== null) pageItem.locked = originalItemLocked;
            } catch (e) { }
            try {
                if (pageItem && originalItemHidden !== null) pageItem.hidden = originalItemHidden;
            } catch (e) { }
            try {
                if (targetLayer && originalLayerLocked !== null) targetLayer.locked = originalLayerLocked;
            } catch (e) { }
            try {
                if (targetLayer && originalLayerVisible !== null) targetLayer.visible = originalLayerVisible;
            } catch (e) { }
        }
    }

    /**
     * テキストの内容を置き換える。段落書式を保持するときは段落単位の書式だけを控えて戻す（文字単位の書式は保持しない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} newText - 新しい内容
     * @param {boolean} keepParagraphFormat - 段落書式を保持するなら true
     * @returns {void}
     */
    function replaceContent(textFrame, newText, keepParagraphFormat) {
        if (!/text/i.test(textFrame.typename)) return;

        withTemporarilyEditableItem(textFrame, function () {
            if (!keepParagraphFormat) {
                textFrame.contents = newText;
                return;
            }

            var paragraphs = getParagraphs(textFrame);
            var paragraphSnapshots = [];

            try {
                for (var i = 0; i < paragraphs.length; i++) {
                    paragraphSnapshots.push(readProperties(paragraphs[i], PARAGRAPH_FORMAT_KEYS));
                }
            } catch (err) { }

            textFrame.contents = newText;
            paragraphs = getParagraphs(textFrame);

            if (paragraphSnapshots.length === 0) {
                return;
            }

            /* 段落が増えたら最後の段落の書式を使う / Extra paragraphs take the last snapshot */
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraphSnapshot = paragraphSnapshots[j] ? paragraphSnapshots[j] : paragraphSnapshots[paragraphSnapshots.length - 1];
                writeProperties(paragraphSnapshot, paragraphs[j], PARAGRAPH_FORMAT_KEYS);
            }
        });
    }

    // =========================================
    // ソート / Sorting
    // =========================================

    /**
     * 上から下へ、ほぼ同じ高さなら左から右へ並べ替える
     * @param {TextFrame[]} textFrames - 並べ替える配列（その場で並べ替える）
     * @param {number} [tolerance] - 同じ高さとみなす差（省略時は POSITION_SORT_TOLERANCE）
     * @returns {void}
     */
    function sortByPosition(textFrames, tolerance) {
        if (!tolerance) tolerance = POSITION_SORT_TOLERANCE;
        textFrames.sort(function (firstFrame, secondFrame) {
            if (Math.abs(secondFrame.top - firstFrame.top) <= tolerance) {
                return firstFrame.left - secondFrame.left;
            }
            return secondFrame.top - firstFrame.top;
        });
    }

    /**
     * 内容の順（大文字小文字を区別しない）に並べ替える
     * @param {TextFrame[]} textFrames - 並べ替える配列（その場で並べ替える）
     * @returns {void}
     */
    function sortByContent(textFrames) {
        textFrames.sort(function (firstFrame, secondFrame) {
            var firstText = firstFrame.contents.toLowerCase();
            var secondText = secondFrame.contents.toLowerCase();
            if (firstText < secondText) return -1;
            if (firstText > secondText) return 1;
            return 0;
        });
    }

    // =========================================
    // シンボル内テキスト収集 / Collect text inside symbols
    // =========================================

    /**
     * 作業レイヤーを探し、無ければ作る
     * @param {Document} doc - 対象のドキュメント
     * @returns {{layer: Layer, created: boolean}} 作業レイヤーと、新しく作ったかどうか
     */
    function getOrCreateTempLayer(doc) {
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
     * 作業レイヤーの中身を空にする
     * @param {Layer} tempLayer - 作業レイヤー
     * @returns {void}
     */
    function clearTempLayer(tempLayer) {
        if (!tempLayer) return;
        try {
            while (tempLayer.pageItems.length > 0) {
                tempLayer.pageItems[0].remove();
            }
        } catch (e) { }
    }

    /**
     * 作業レイヤーを空にして削除する
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function removeTempLayer(doc) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === TEMP_LAYER_NAME && doc.layers[i].note === TEMP_LAYER_NOTE) {
                clearTempLayer(doc.layers[i]);
                try {
                    doc.layers[i].remove();
                } catch (e2) { }
                break;
            }
        }
    }

    /**
     * シンボルインスタンスを作業レイヤーの先頭へ複製する（選択は外す）
     * @param {SymbolItem} symbolItem - 元のシンボルインスタンス
     * @param {Layer} targetLayer - 作業レイヤー
     * @returns {SymbolItem} 複製
     */
    function duplicateSymbolItemToLayer(symbolItem, targetLayer) {
        var duplicateItem = symbolItem.duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING);
        duplicateItem.selected = false;
        return duplicateItem;
    }

    /**
     * オブジェクト（グループなどは中まで）から空でないテキストフレームを集める
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームを入れる配列
     * @returns {void}
     */
    function collectTextFramesFromItem(pageItem, collectedFrames) {
        if (!pageItem) return;
        if (!collectedFrames) return;

        if (pageItem.typename === 'TextFrame') {
            if (!isEmptyTextFrame(pageItem)) {
                collectedFrames.push(pageItem);
            }
            return;
        }

        if (!pageItem.pageItems) return;
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            collectTextFramesFromItem(pageItem.pageItems[i], collectedFrames);
        }
    }

    /**
     * オブジェクト群に含まれるテキストの内容を集める
     * @param {PageItem[]} pageItems - 対象のオブジェクト
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
     * オブジェクト群を削除する（削除できないものは飛ばす）
     * @param {PageItem[]} pageItems - 削除するオブジェクト
     * @returns {void}
     */
    function removeItems(pageItems) {
        for (var i = pageItems.length - 1; i >= 0; i--) {
            try {
                if (pageItems[i] && pageItems[i].isValid !== false) {
                    pageItems[i].remove();
                }
            } catch (e) { }
        }
    }

    /**
     * 処理のあと、元の選択に戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {Function} selectionAction - 選択を変える処理
     * @returns {void}
     */
    function withSelectionRestored(doc, selectionAction) {
        var previousSelection = [];
        try {
            for (var i = 0; i < doc.selection.length; i++) {
                previousSelection.push(doc.selection[i]);
            }
        } catch (e) { }

        try {
            selectionAction();
        } finally {
            try {
                doc.selection = null;
            } catch (e2) { }
            for (var j = 0; j < previousSelection.length; j++) {
                try {
                    previousSelection[j].selected = true;
                } catch (e3) { }
            }
        }
    }

    /**
     * シンボルインスタンスが範囲に入るか
     * @param {Document} doc - 対象のドキュメント
     * @param {SymbolItem} symbolItem - 調べるシンボルインスタンス
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @param {number} activeArtboardIndex - 現在のアートボードの番号
     * @returns {boolean} 範囲に入れば true
     */
    function isSymbolItemInScope(doc, symbolItem, scopeMode, activeArtboardIndex) {
        var artboardIndex = getItemArtboardIndex(doc, symbolItem);
        if (scopeMode === 'current') {
            return artboardIndex === activeArtboardIndex;
        }
        if (scopeMode === 'allArtboards') {
            return artboardIndex >= 0;
        }
        return true;
    }

    /**
     * 範囲に入るシンボルインスタンスを、シンボル名とアートボードの組ごとに1つずつ集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { item, symbolName, artboardIndex } の配列
     */
    function collectScopedSymbolItems(doc, collectOptions, scopeMode) {
        var symbolItems = doc.symbolItems;
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var scopedSymbolItems = [];
        var processedKeys = {};

        for (var i = 0; i < symbolItems.length; i++) {
            var symbolItem = symbolItems[i];
            var symbolName = '';

            if (!symbolItem || symbolItem.isValid === false) continue;
            if (!collectOptions.includeComment && isCommentLayer(symbolItem)) continue;
            if (!isCollectable(symbolItem, collectOptions)) continue;
            if (!isSymbolItemInScope(doc, symbolItem, scopeMode, activeArtboardIndex)) continue;

            try {
                symbolName = symbolItem.symbol.name;
            } catch (e) { }
            if (!symbolName) continue;

            var artboardIndex = getItemArtboardIndex(doc, symbolItem);
            var processKey = symbolName + '||' + artboardIndex;
            if (processedKeys[processKey]) continue;

            processedKeys[processKey] = true;
            scopedSymbolItems.push({
                item: symbolItem,
                symbolName: symbolName,
                artboardIndex: artboardIndex
            });
        }

        return scopedSymbolItems;
    }

    /**
     * シンボルインスタンスを作業レイヤーに複製してリンクを解除し、中のテキストを1件ずつ渡す
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} scopedItems - collectScopedSymbolItems() の結果
     * @param {Function} onText - (symbolName, artboardIndex, text) を受け取る関数
     * @returns {void}
     */
    function collectTextsFromScopedSymbolItems(doc, scopedItems, onText) {
        if (!scopedItems || scopedItems.length === 0) {
            return;
        }

        var tempInfo = getOrCreateTempLayer(doc);
        var tempLayer = tempInfo.layer;
        clearTempLayer(tempLayer);

        withSelectionRestored(doc, function () {
            for (var index = scopedItems.length - 1; index >= 0; index--) {
                var scopedEntry = scopedItems[index];
                var symbolItem = scopedEntry.item;
                var brokenItems = [];

                if (!symbolItem || symbolItem.isValid === false) continue;

                try {
                    var workingSymbolItem = duplicateSymbolItemToLayer(symbolItem, tempLayer);
                    try {
                        doc.selection = null;
                    } catch (e2) { }
                    workingSymbolItem.selected = true;
                    workingSymbolItem.breakLink();

                    /* リンク解除でできたものは選択に残る / breakLink leaves its result selected */
                    try {
                        for (var s = 0; s < doc.selection.length; s++) {
                            brokenItems.push(doc.selection[s]);
                        }
                    } catch (e3) { }

                    if (brokenItems.length > 0) {
                        var textContents = extractTextContentsFromItems(brokenItems);
                        for (var t = 0; t < textContents.length; t++) {
                            onText(scopedEntry.symbolName, scopedEntry.artboardIndex, textContents[t]);
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
            removeTempLayer(doc);
        }
        try {
            doc.selection = null;
        } catch (e6) { }
    }

    /**
     * シンボル内テキスト一覧用に「シンボル名：内容」を集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {string[]} 一覧の行
     */
    function collectSymbolTexts(doc, collectOptions, scopeMode) {
        var symbolLines = [];
        var scopedItems = collectScopedSymbolItems(doc, collectOptions, scopeMode);

        collectTextsFromScopedSymbolItems(doc, scopedItems, function (symbolName, artboardIndex, text) {
            symbolLines.push(symbolName + getLabel('format.symbolSeparator') + text.replace(/[\r\n]+/g, ' '));
        });

        return symbolLines;
    }

    /**
     * 書き出し用に、シンボル内テキストを「内容〈シンボル：名前〉」の形でアートボード番号つきで集める
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { artboardIndex, text } の配列
     */
    function collectSymbolTextsByArtboard(doc, collectOptions, scopeMode) {
        var symbolEntries = [];
        var scopedItems = collectScopedSymbolItems(doc, collectOptions, scopeMode);

        collectTextsFromScopedSymbolItems(doc, scopedItems, function (symbolName, artboardIndex, text) {
            symbolEntries.push({
                artboardIndex: artboardIndex,
                text: text + getLabel('format.symbolSuffix').split('{symbolName}').join(symbolName)
            });
        });

        return symbolEntries;
    }

    // =========================================
    // テキスト書き出し / Text export
    // =========================================

    /**
     * 範囲に応じたアートボードごとの入れ物を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object} { groups, groupIndexMap, activeArtboardIndex, outsideGroup }（各グループは { name, items }）
     */
    function createArtboardGroups(doc, scopeMode) {
        var artboardGroups = {
            groups: [],
            groupIndexMap: {},
            activeArtboardIndex: doc.artboards.getActiveArtboardIndex(),
            outsideGroup: { name: getLabel('fallbackName.outsideArtboards'), items: [] }
        };

        if (scopeMode === 'current') {
            artboardGroups.groups.push({
                name: getArtboardDisplayName(doc, artboardGroups.activeArtboardIndex),
                items: []
            });
            artboardGroups.groupIndexMap[artboardGroups.activeArtboardIndex] = 0;
        } else {
            for (var i = 0; i < doc.artboards.length; i++) {
                artboardGroups.groupIndexMap[i] = artboardGroups.groups.length;
                artboardGroups.groups.push({
                    name: getArtboardDisplayName(doc, i),
                    items: []
                });
            }
        }
        return artboardGroups;
    }

    /**
     * 値をアートボードのグループへ振り分ける（範囲外は捨て、ドキュメント全体ならアートボード外へ）
     * @param {Object} artboardGroups - createArtboardGroups() の結果
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @param {number} artboardIndex - 値のあるアートボードの番号（どこにも無ければ -1）
     * @param {*} itemValue - 振り分ける値
     * @returns {void}
     */
    function addToArtboardGroup(artboardGroups, scopeMode, artboardIndex, itemValue) {
        if (scopeMode === 'current') {
            if (artboardIndex === artboardGroups.activeArtboardIndex) {
                artboardGroups.groups[artboardGroups.groupIndexMap[artboardIndex]].items.push(itemValue);
            }
        } else if (artboardIndex >= 0) {
            artboardGroups.groups[artboardGroups.groupIndexMap[artboardIndex]].items.push(itemValue);
        } else if (scopeMode !== 'allArtboards') {
            artboardGroups.outsideGroup.items.push(itemValue);
        }
    }

    /**
     * アートボードのグループを確定する（ドキュメント全体でアートボード外があれば末尾に加える）
     * @param {Object} artboardGroups - createArtboardGroups() の結果
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { name, items } の配列
     */
    function finishArtboardGroups(artboardGroups, scopeMode) {
        if (scopeMode === 'all' && artboardGroups.outsideGroup.items.length > 0) {
            artboardGroups.groups.push(artboardGroups.outsideGroup);
        }
        return artboardGroups.groups;
    }

    /**
     * 書き出すテキスト（テキストフレームとシンボル内テキスト）をアートボードごとにまとめる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { name, items（テキストの内容） } の配列
     */
    function collectArtboardGroupedExportData(doc, collectOptions, scopeMode) {
        var textFrames = collectFramesByScope(doc, scopeMode, collectOptions);
        var symbolEntries = collectSymbolTextsByArtboard(doc, collectOptions, scopeMode);
        var artboardGroups = createArtboardGroups(doc, scopeMode);
        var i;

        for (i = 0; i < textFrames.length; i++) {
            addToArtboardGroup(artboardGroups, scopeMode, getItemArtboardIndex(doc, textFrames[i]), textFrames[i].contents);
        }
        for (i = 0; i < symbolEntries.length; i++) {
            addToArtboardGroup(artboardGroups, scopeMode, symbolEntries[i].artboardIndex, symbolEntries[i].text);
        }
        return finishArtboardGroups(artboardGroups, scopeMode);
    }

    /**
     * フォント名を書き出すためのテキストフレームをアートボードごとにまとめる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {Object[]} { name, items（テキストフレーム） } の配列
     */
    function collectArtboardGroupedFontFrames(doc, collectOptions, scopeMode) {
        var textFrames = collectFramesByScope(doc, scopeMode, collectOptions);
        var artboardGroups = createArtboardGroups(doc, scopeMode);

        for (var i = 0; i < textFrames.length; i++) {
            addToArtboardGroup(artboardGroups, scopeMode, getItemArtboardIndex(doc, textFrames[i]), textFrames[i]);
        }
        return finishArtboardGroups(artboardGroups, scopeMode);
    }

    /**
     * 書き出すテキストをアートボードごとに組み立てる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} exportSettings - { includeText, includeFonts, openAfter }
     * @param {Object} collectOptions - { includeComment, includeLocked, includeHidden }
     * @param {string} scopeMode - "all" / "allArtboards" / "current"
     * @returns {string} 書き出す内容
     */
    function buildExportText(doc, exportSettings, collectOptions, scopeMode) {
        var artboardGroups = collectArtboardGroupedExportData(doc, collectOptions, scopeMode);
        var artboardFontGroups = collectArtboardGroupedFontFrames(doc, collectOptions, scopeMode);
        var exportLines = [];

        for (var i = 0; i < artboardGroups.length; i++) {
            if (i > 0) exportLines.push('');
            exportLines.push('---' + artboardGroups[i].name + '---');

            if (exportSettings.includeText) {
                exportLines.push('[Text]');
                for (var j = 0; j < artboardGroups[i].items.length; j++) {
                    exportLines.push(artboardGroups[i].items[j]);
                }
                exportLines.push('');
            }

            if (exportSettings.includeFonts) {
                exportLines.push('[Font Names]');
                var fontLines = collectFontDisplayNamesFromFrames(artboardFontGroups[i].items);
                for (var k = 0; k < fontLines.length; k++) {
                    exportLines.push(fontLines[k]);
                }
                exportLines.push('');
            }
        }

        return exportLines.join('\n');
    }

    /**
     * 書き出しオプションのダイアログを表示する
     * @returns {Object|null} OK なら { includeText, includeFonts, openAfter }、キャンセルなら null
     */
    function showExportOptionsDialog() {
        var exportDialog = new Window('dialog', getLabel('dialog.exportOptions'));
        exportDialog.orientation = 'column';
        exportDialog.alignChildren = ['fill', 'top'];

        var optionsPanel = exportDialog.add('panel', undefined, getLabel('dialog.exportOptions'));
        optionsPanel.orientation = 'column';
        optionsPanel.alignChildren = ['left', 'top'];
        optionsPanel.margins = EXPORT_PANEL_MARGINS;

        var cbExportText = optionsPanel.add('checkbox', undefined, getLabel('checkbox.exportIncludeText'));
        cbExportText.value = true;

        var cbExportFonts = optionsPanel.add('checkbox', undefined, getLabel('checkbox.exportIncludeFonts'));
        cbExportFonts.value = true;

        /* 「書き出し後にファイルを開く」はパネルの外に置く / "Open file after export" sits below the panel */
        var cbOpenAfter = exportDialog.add('checkbox', undefined, getLabel('checkbox.exportOpenAfter'));
        cbOpenAfter.value = true;

        var btnRowGroup = exportDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['center', 'center'];
        btnRowGroup.alignment = ['center', 'center'];

        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        exportDialog.defaultElement = btnOK;

        if (exportDialog.show() !== 1) {
            return null;
        }

        return {
            includeText: cbExportText.value,
            includeFonts: cbExportFonts.value,
            openAfter: cbOpenAfter.value
        };
    }

    /**
     * 書き出すファイルのパス（デスクトップの text-<ドキュメント名>-<日時>.txt）を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {string} ファイルのパス
     */
    function buildExportFilePath(doc) {
        return Folder.desktop.fsName + '/text-' + sanitizeFileName(getDocumentBaseName(doc)) + '-' + getDateTimeStamp() + '.txt';
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 右カラムのパネル（縦並び・左揃え）を追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - パネル名の LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addScopePanel(parentGroup, labelPath) {
        var scopePanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        scopePanel.orientation = "column";
        scopePanel.alignChildren = ["left", "top"];
        scopePanel.margins = SCOPE_PANEL_MARGINS;
        return scopePanel;
    }

    /**
     * 名前一覧用のタブ（見出し＋横並びの切り替え）を追加する
     * @param {TabbedPanel} infoTabs - 追加先
     * @param {string} tabLabelPath - タブ名の LABELS パス
     * @param {string} headingLabelPath - 見出しの LABELS パス
     * @returns {{tab: Tab, choiceRow: Group}} 追加したタブと、切り替えを並べる行
     */
    function addInfoTab(infoTabs, tabLabelPath, headingLabelPath) {
        var infoTab = infoTabs.add("tab", undefined, getLabel(tabLabelPath));
        infoTab.orientation = "column";
        infoTab.alignChildren = ["fill", "top"];
        infoTab.margins = INFO_TAB_MARGINS;
        infoTab.spacing = INFO_TAB_SPACING;

        infoTab.add("statictext", undefined, getLabel(headingLabelPath));

        var choiceRow = infoTab.add("group");
        choiceRow.orientation = "row";
        choiceRow.alignChildren = ["left", "center"];
        choiceRow.spacing = CHOICE_ROW_SPACING;
        return { tab: infoTab, choiceRow: choiceRow };
    }

    /**
     * 読み取り専用の複数行テキスト欄を追加する
     * @param {Tab} parentTab - 追加先
     * @returns {EditText} 追加した欄
     */
    function addReadOnlyTextArea(parentTab) {
        var textArea = parentTab.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: true });
        textArea.alignment = ["fill", "fill"];
        return textArea;
    }

    /**
     * ［カンバス］タブの左カラム（テキスト一覧・編集欄・シンボル内テキスト）を組む
     * @param {Group} mainGroup - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildTextColumn(mainGroup, dialogControls) {
        var textColumn = mainGroup.add("group");
        textColumn.orientation = "column";
        textColumn.alignChildren = ["fill", "fill"];

        textColumn.add("statictext", undefined, getLabel("fieldLabel.textList"));
        dialogControls.textListBox = textColumn.add("listbox", TEXT_LIST_BOUNDS, []);
        textColumn.add("statictext", undefined, getLabel("fieldLabel.textEdit"));
        dialogControls.textEditBox = textColumn.add("edittext", TEXT_EDIT_BOUNDS, "", { multiline: true, scrolling: true });
        dialogControls.textEditBox.helpTip = getLabel("tooltip.textEdit").replace("{softBreak}", SOFT_BREAK);

        textColumn.add("statictext", undefined, getLabel("fieldLabel.symbolText"));
        dialogControls.symbolListBox = textColumn.add("listbox", [0, 0, SYMBOL_LIST_WIDTH, SYMBOL_LIST_MIN_ROWS * SYMBOL_LIST_ROW_HEIGHT + SYMBOL_LIST_EXTRA_HEIGHT], []);
        dialogControls.symbolListBox.alignment = ["fill", "fill"];
    }

    /**
     * ［カンバス］タブの右カラム（対象・レイヤー・ソート・オプション）を組む
     * @param {Group} mainGroup - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildScopeColumn(mainGroup, dialogControls) {
        var scopeColumn = mainGroup.add("group");
        scopeColumn.orientation = "column";
        scopeColumn.alignChildren = ["fill", "top"];

        /* 対象テキストパネル / Target text panel */
        var targetPanel = addScopePanel(scopeColumn, "panel.targetText");
        var artboardRadioGroup = targetPanel.add("group");
        artboardRadioGroup.orientation = "column";
        artboardRadioGroup.alignChildren = ["left", "top"];
        dialogControls.rbCurrentArtboard = artboardRadioGroup.add("radiobutton", undefined, getLabel("radio.currentArtboard"));
        dialogControls.rbAllArtboards = artboardRadioGroup.add("radiobutton", undefined, getLabel("radio.allArtboards"));
        dialogControls.rbCurrentArtboard.value = true;

        var wholeDocumentRow = targetPanel.add("group");
        wholeDocumentRow.orientation = "row";
        dialogControls.cbWholeDocument = wholeDocumentRow.add("checkbox", undefined, getLabel("checkbox.wholeDocument"));
        dialogControls.cbWholeDocument.value = false;
        dialogControls.cbWholeDocument.enabled = false;
        dialogControls.cbWholeDocument.helpTip = getLabel("tooltip.wholeDocument");

        /* レイヤーなどの条件 / Layer conditions */
        var layerPanel = addScopePanel(scopeColumn, "panel.layerText");
        dialogControls.cbIncludeCommentLayers = layerPanel.add("checkbox", undefined, getLabel("checkbox.includeCommentLayers"));
        dialogControls.cbIncludeCommentLayers.value = false;
        dialogControls.cbIncludeLocked = layerPanel.add("checkbox", undefined, getLabel("checkbox.includeLocked"));
        dialogControls.cbIncludeLocked.value = false;
        dialogControls.cbIncludeLocked.helpTip = getLabel("tooltip.includeLocked");
        dialogControls.cbIncludeHidden = layerPanel.add("checkbox", undefined, getLabel("checkbox.includeHidden"));
        dialogControls.cbIncludeHidden.value = false;
        dialogControls.cbIncludeHidden.helpTip = getLabel("tooltip.includeHidden");

        /* ソート / Sort */
        var sortPanel = addScopePanel(scopeColumn, "panel.sort");
        var sortRadioGroup = sortPanel.add("group");
        sortRadioGroup.orientation = "row";
        sortRadioGroup.alignChildren = ["left", "center"];
        dialogControls.rbSortNone = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortNone"));
        dialogControls.rbSortPosition = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortPosition"));
        dialogControls.rbSortPosition.helpTip = getLabel("tooltip.sortPosition");
        dialogControls.rbSortAlphabetical = sortRadioGroup.add("radiobutton", undefined, getLabel("radio.sortAlphabetical"));
        dialogControls.rbSortAlphabetical.helpTip = getLabel("tooltip.sortAlphabetical");
        dialogControls.rbSortNone.value = true;

        /* オプション / Options */
        var optionsGroup = scopeColumn.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "top"];
        optionsGroup.margins = OPTIONS_GROUP_MARGINS;

        dialogControls.cbMergeDuplicates = optionsGroup.add("checkbox", undefined, getLabel("checkbox.mergeDuplicates"));
        dialogControls.cbMergeDuplicates.value = true;
        dialogControls.cbMergeDuplicates.helpTip = getLabel("tooltip.mergeDuplicates");

        dialogControls.cbKeepFormat = optionsGroup.add("checkbox", undefined, getLabel("checkbox.keepFormat"));
        dialogControls.cbKeepFormat.value = true;
        dialogControls.cbKeepFormat.helpTip = getLabel("tooltip.keepFormat");

        dialogControls.cbPreview = optionsGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        dialogControls.cbPreview.helpTip = getLabel("tooltip.preview");
    }

    /**
     * ［カンバス］タブを組む
     * @param {TabbedPanel} infoTabs - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {Tab} ［カンバス］タブ
     */
    function buildCanvasTab(infoTabs, dialogControls) {
        var canvasTab = infoTabs.add("tab", undefined, getLabel("tab.canvas"));
        canvasTab.orientation = "row";
        canvasTab.alignChildren = ["fill", "fill"];
        canvasTab.margins = CANVAS_TAB_MARGINS;
        canvasTab.spacing = CANVAS_TAB_SPACING;

        var mainGroup = canvasTab.add("group");
        mainGroup.orientation = "row";
        mainGroup.alignChildren = ["fill", "fill"];

        buildTextColumn(mainGroup, dialogControls);
        buildScopeColumn(mainGroup, dialogControls);
        return canvasTab;
    }

    /**
     * ［レイヤー名］［アートボード名］［フォント名］タブを組む
     * @param {TabbedPanel} infoTabs - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildInfoTabs(infoTabs, dialogControls) {
        var layerNamesTab = addInfoTab(infoTabs, "tab.layerNames", "fieldLabel.layerNameList");
        dialogControls.rbLayerScopeTop = layerNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.layerScopeTop"));
        dialogControls.rbLayerScopeAll = layerNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.layerScopeAll"));
        dialogControls.rbLayerScopeTop.value = true;
        dialogControls.layerNameEdit = addReadOnlyTextArea(layerNamesTab.tab);

        var artboardNamesTab = addInfoTab(infoTabs, "tab.artboardNames", "fieldLabel.artboardNameList");
        dialogControls.rbArtboardScopeNumbered = artboardNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.artboardScopeNumbered"));
        dialogControls.rbArtboardScopeRaw = artboardNamesTab.choiceRow.add("radiobutton", undefined, getLabel("radio.artboardScopeRaw"));
        dialogControls.rbArtboardScopeNumbered.value = true;
        dialogControls.artboardNameEdit = addReadOnlyTextArea(artboardNamesTab.tab);

        var fontNamesTab = addInfoTab(infoTabs, "tab.fontNames", "fieldLabel.fontList");
        dialogControls.cbFontPS = fontNamesTab.choiceRow.add("checkbox", undefined, getLabel("checkbox.fontPS"));
        dialogControls.cbFontPS.value = true;
        dialogControls.cbFontFamily = fontNamesTab.choiceRow.add("checkbox", undefined, getLabel("checkbox.fontFamily"));
        dialogControls.cbFontFamily.value = true;
        dialogControls.cbFontStyle = fontNamesTab.choiceRow.add("checkbox", undefined, getLabel("checkbox.fontStyle"));
        dialogControls.cbFontStyle.value = true;

        dialogControls.fontNameListBox = fontNamesTab.tab.add("listbox", undefined, [], {
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [getLabel("listColumn.fontPS"), getLabel("listColumn.fontFamily"), getLabel("listColumn.fontStyle")],
            columnWidths: FONT_COLUMN_WIDTHS
        });
        dialogControls.fontNameListBox.alignment = ["fill", "fill"];
        dialogControls.fontNameListBox.helpTip = getLabel("tooltip.fontList");
    }

    /**
     * ボタン行（左に書き出し、右にキャンセル・OK）を組む
     * @param {Window} textScopeDialog - 追加先のダイアログ
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildButtonRow(textScopeDialog, dialogControls) {
        var btnRowGroup = textScopeDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        dialogControls.btnExportText = btnLeftGroup.add("button", undefined, getLabel("button.exportText"));
        dialogControls.btnExportText.helpTip = getLabel("tooltip.exportText");

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        dialogControls.btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        textScopeDialog.defaultElement = dialogControls.btnOK;
    }

    /**
     * CC 2020（v24）かどうか。プレビューでクラッシュするため無効にする
     * @returns {boolean} v24 なら true
     */
    function isPreviewUnsupported() {
        return parseInt(app.version) == 24;
    }

    /**
     * メインダイアログを組み立てる（イベントは未接続）
     * @returns {Object} コントロールをまとめたオブジェクト
     */
    function buildDialog() {
        var dialogControls = {};
        var textScopeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        textScopeDialog.orientation = "column";
        textScopeDialog.alignChildren = ["fill", "top"];
        dialogControls.textScopeDialog = textScopeDialog;

        /* ダイアログ全体をタブで構成 / Use tabs for the whole dialog */
        var infoTabs = textScopeDialog.add("tabbedpanel");
        infoTabs.alignChildren = ["fill", "fill"];
        infoTabs.preferredSize = TABS_SIZE;
        infoTabs.margins = TABS_MARGINS;

        var canvasTab = buildCanvasTab(infoTabs, dialogControls);
        buildInfoTabs(infoTabs, dialogControls);
        infoTabs.selection = canvasTab;

        buildButtonRow(textScopeDialog, dialogControls);

        /* CC 2020 v24.3 はプレビュー時にクラッシュするため無効化 / Disable preview in CC 2020 v24.3 because it may crash */
        if (isPreviewUnsupported()) {
            dialogControls.cbPreview.enabled = false;
            dialogControls.cbPreview.helpTip = getLabel("tooltip.previewDisabled");
        }

        return dialogControls;
    }

    // =========================================
    // 一覧の更新 / Updating the lists
    // =========================================

    /**
     * 対象テキストの範囲を返す
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {string} "all" / "allArtboards" / "current"
     */
    function getCurrentScopeMode(dialogControls) {
        if (dialogControls.cbWholeDocument.value) {
            return 'all';
        }
        if (dialogControls.rbAllArtboards.value) {
            return 'allArtboards';
        }
        return 'current';
    }

    /**
     * レイヤーなどの条件をダイアログから読む
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {Object} { includeComment, includeLocked, includeHidden }
     */
    function readCollectOptions(dialogControls) {
        return {
            includeComment: dialogControls.cbIncludeCommentLayers.value,
            includeLocked: dialogControls.cbIncludeLocked.value,
            includeHidden: dialogControls.cbIncludeHidden.value
        };
    }

    /**
     * シンボル内テキストのキャッシュのキー（範囲と条件の組み合わせ）を返す
     * @param {Object} editSession - 編集の状態
     * @returns {string} キャッシュのキー
     */
    function getCurrentSymbolCacheKey(editSession) {
        var dialogControls = editSession.dialogControls;
        var scopeMode = getCurrentScopeMode(dialogControls);

        var abIndexPart = '';
        if (scopeMode === 'current') {
            try {
                abIndexPart = 'ab' + editSession.doc.artboards.getActiveArtboardIndex();
            } catch (e) {
                abIndexPart = 'ab0';
            }
        }

        return [
            scopeMode,
            abIndexPart,
            dialogControls.cbIncludeCommentLayers.value ? 'comment1' : 'comment0',
            dialogControls.cbIncludeLocked.value ? 'locked1' : 'locked0',
            dialogControls.cbIncludeHidden.value ? 'hidden1' : 'hidden0'
        ].join('|');
    }

    /**
     * ［レイヤー名］［アートボード名］［フォント名］タブの表示を更新する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function refreshInfoTabs(editSession) {
        var doc = editSession.doc;
        var dialogControls = editSession.dialogControls;
        var textFrames = getScopedFramesForInfoTabs(doc);
        var layerNames = dialogControls.rbLayerScopeTop.value ? collectTopLevelLayerNames(doc) : collectAllDocumentLayerNames(doc);
        var artboardNames = [];
        var useRawName = dialogControls.rbArtboardScopeRaw.value;
        var i;

        for (i = 0; i < doc.artboards.length; i++) {
            artboardNames.push(useRawName ? getRawArtboardName(doc, i) : getArtboardTabDisplayName(doc, i));
        }

        var fontRows = collectUniqueFontTriples(textFrames);

        layerNames.sort();
        artboardNames.sort();
        fontRows.sort(function (firstRow, secondRow) {
            var firstKey = getFontTripleKey(firstRow);
            var secondKey = getFontTripleKey(secondRow);
            if (firstKey < secondKey) return -1;
            if (firstKey > secondKey) return 1;
            return 0;
        });

        dialogControls.layerNameEdit.text = layerNames.join("\n");
        dialogControls.artboardNameEdit.text = artboardNames.join("\n");
        dialogControls.fontNameListBox.removeAll();
        for (i = 0; i < fontRows.length; i++) {
            var fontListRow = dialogControls.fontNameListBox.add('item', dialogControls.cbFontPS.value ? fontRows[i].psName : '');
            fontListRow.subItems[0].text = dialogControls.cbFontFamily.value ? fontRows[i].familyName : '';
            fontListRow.subItems[1].text = dialogControls.cbFontStyle.value ? fontRows[i].styleName : '';
            fontListRow.psNameKey = fontRows[i].psName;
        }
    }

    /**
     * テキスト一覧を集め直して表示する（プレビューは先に戻す）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function updateList(editSession) {
        var dialogControls = editSession.dialogControls;
        clearPreviewIfNeeded(editSession);

        var scopeMode = getCurrentScopeMode(dialogControls);
        var collectOptions = readCollectOptions(dialogControls);
        editSession.duplicateMap = [];
        editSession.textFrameList = collectFramesByScope(editSession.doc, scopeMode, collectOptions);

        /* ソート / Sorting */
        if (dialogControls.rbSortPosition.value) {
            sortByPosition(editSession.textFrameList);
        } else if (dialogControls.rbSortAlphabetical.value) {
            sortByContent(editSession.textFrameList);
        }

        if (dialogControls.cbMergeDuplicates.value) {
            var duplicateGroups = groupDuplicateFrames(editSession.textFrameList);
            editSession.duplicateMap = duplicateGroups.duplicateMap;
            editSession.textFrameList = duplicateGroups.uniqueFrames;
        }

        var textListBox = dialogControls.textListBox;
        textListBox.removeAll();
        for (var i = 0; i < editSession.textFrameList.length; i++) {
            textListBox.add("item", (i + 1) + getLabel("format.itemPrefix") + buildListLabel(editSession.textFrameList[i].contents));
        }
        dialogControls.textEditBox.text = "";
        if (editSession.textFrameList.length > 0) {
            textListBox.selection = 0;
        }
        refreshInfoTabs(editSession);
    }

    /**
     * シンボル内テキストを返す（範囲と条件ごとにキャッシュ）
     * @param {Object} editSession - 編集の状態
     * @param {boolean} forceRefresh - キャッシュを使わずに集め直すなら true
     * @returns {string[]} 一覧の行
     */
    function getSymbolTexts(editSession, forceRefresh) {
        var symbolTextCache = editSession.symbolTextCache;
        var cacheKey = getCurrentSymbolCacheKey(editSession);
        if (forceRefresh) {
            delete symbolTextCache[cacheKey];
        }
        if (symbolTextCache.hasOwnProperty(cacheKey)) {
            return symbolTextCache[cacheKey];
        }
        var dialogControls = editSession.dialogControls;
        symbolTextCache[cacheKey] = collectSymbolTexts(editSession.doc, readCollectOptions(dialogControls), getCurrentScopeMode(dialogControls));
        return symbolTextCache[cacheKey];
    }

    /**
     * 行数に合わせてシンボル内テキスト一覧の高さを変える
     * @param {Object} dialogControls - buildDialog() の結果
     * @param {number} itemCount - 行数
     * @returns {void}
     */
    function updateSymbolListHeight(dialogControls, itemCount) {
        var visibleRows = itemCount;
        if (visibleRows < SYMBOL_LIST_MIN_ROWS) visibleRows = SYMBOL_LIST_MIN_ROWS;
        if (visibleRows > SYMBOL_LIST_MAX_ROWS) visibleRows = SYMBOL_LIST_MAX_ROWS;
        dialogControls.symbolListBox.preferredSize.height = visibleRows * SYMBOL_LIST_ROW_HEIGHT + SYMBOL_LIST_EXTRA_HEIGHT;
        try {
            dialogControls.textScopeDialog.layout.layout(true);
            dialogControls.textScopeDialog.layout.resize();
        } catch (e) { }
    }

    /**
     * シンボル内テキスト一覧を更新する
     * @param {Object} editSession - 編集の状態
     * @param {boolean} forceRefresh - キャッシュを使わずに集め直すなら true
     * @returns {void}
     */
    function refreshSymbolList(editSession, forceRefresh) {
        var symbolListBox = editSession.dialogControls.symbolListBox;
        symbolListBox.removeAll();
        var symbolLines = getSymbolTexts(editSession, forceRefresh);
        for (var i = 0; i < symbolLines.length; i++) {
            symbolListBox.add("item", symbolLines[i]);
        }
        updateSymbolListHeight(editSession.dialogControls, symbolLines.length);
    }

    // =========================================
    // 編集とプレビュー / Editing and preview
    // =========================================

    /**
     * 選択中の行のテキストに編集を反映する（まとめているときは同じ内容の全フレームへ）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function applyCurrentEdit(editSession) {
        var dialogControls = editSession.dialogControls;
        if (dialogControls.textListBox.selection === null) return;
        var selectedIndex = dialogControls.textListBox.selection.index;
        var newText = dialogControls.textEditBox.text.replace(new RegExp(SOFT_BREAK, 'gmi'), '\x03');
        var keepParagraphFormat = dialogControls.cbKeepFormat.value;
        if (dialogControls.cbMergeDuplicates.value && editSession.duplicateMap[selectedIndex]) {
            var sameFrames = editSession.duplicateMap[selectedIndex];
            for (var i = 0; i < sameFrames.length; i++) {
                replaceContent(sameFrames[i], newText, keepParagraphFormat);
            }
        } else {
            replaceContent(editSession.textFrameList[selectedIndex], newText, keepParagraphFormat);
        }
    }

    /**
     * プレビューを掛け直す、またはプレビューをやめて戻す（app.undo() で1段戻してから適用）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function updatePreview(editSession) {
        if (isPreviewUnsupported()) return;
        var dialogControls = editSession.dialogControls;
        try {
            if (dialogControls.cbPreview.enabled && dialogControls.cbPreview.value && dialogControls.textListBox.selection !== null) {
                if (editSession.isPreviewApplied) app.undo();
                else editSession.isPreviewApplied = true;
                applyCurrentEdit(editSession);
                app.redraw();
            } else if (editSession.isPreviewApplied) {
                app.undo();
                app.redraw();
                editSession.isPreviewApplied = false;
            }
        } catch (err) { }
    }

    /**
     * プレビュー中なら戻す
     * @param {Object} editSession - 編集の状態
     * @returns {boolean} 戻した場合は true
     */
    function clearPreviewIfNeeded(editSession) {
        if (!editSession.isPreviewApplied) return false;
        try {
            app.undo();
            app.redraw();
        } catch (err) { }
        editSession.isPreviewApplied = false;
        return true;
    }

    /**
     * 戻したプレビューを掛け直す
     * @param {Object} editSession - 編集の状態
     * @param {boolean} wasPreviewActive - 戻す前にプレビュー中だったか
     * @returns {void}
     */
    function restorePreviewIfNeeded(editSession, wasPreviewActive) {
        if (!wasPreviewActive) return;
        if (isPreviewUnsupported()) return;
        var cbPreview = editSession.dialogControls.cbPreview;
        if (!cbPreview.enabled || !cbPreview.value) return;
        updatePreview(editSession);
    }

    /**
     * 書き出しオプションを尋ね、デスクトップにテキストを書き出す（プレビューはいったん戻して掛け直す）
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function exportTextToDesktop(editSession) {
        var doc = editSession.doc;
        var dialogControls = editSession.dialogControls;
        var exportFile = null;
        var wasPreviewActive = editSession.isPreviewApplied;
        try {
            clearPreviewIfNeeded(editSession);
            var exportSettings = showExportOptionsDialog();
            if (!exportSettings) {
                restorePreviewIfNeeded(editSession, wasPreviewActive);
                return;
            }
            var exportContent = buildExportText(doc, exportSettings, readCollectOptions(dialogControls), getCurrentScopeMode(dialogControls));
            var filePath = buildExportFilePath(doc);
            exportFile = new File(filePath);
            exportFile.encoding = 'UTF-8';
            exportFile.lineFeed = 'Unix';
            if (!exportFile.open('w')) {
                throw new Error('open failed: ' + filePath);
            }
            exportFile.write(exportContent);
            exportFile.close();
            restorePreviewIfNeeded(editSession, wasPreviewActive);
            if (exportSettings.openAfter) {
                try {
                    exportFile.execute();
                } catch (openErr) { }
            }
        } catch (e) {
            try {
                if (exportFile && exportFile.opened) exportFile.close();
            } catch (closeErr) { }
            restorePreviewIfNeeded(editSession, wasPreviewActive);
            alert(getLabel('alert.exportFailed') + '\n' + e);
        }
    }

    // =========================================
    // イベントハンドラ / Event handlers
    // =========================================

    /**
     * ダイアログのイベントを接続する
     * @param {Object} editSession - 編集の状態
     * @returns {void}
     */
    function bindDialogEvents(editSession) {
        var dialogControls = editSession.dialogControls;

        /* 一覧の選択で編集欄を更新 / Update the edit box when the list selection changes */
        dialogControls.textListBox.onChange = function () {
            if (dialogControls.textListBox.selection !== null) {
                var selectedIndex = dialogControls.textListBox.selection.index;
                dialogControls.textEditBox.text = editSession.textFrameList[selectedIndex].contents.replace(/\x03/g, SOFT_BREAK);
            }
        };

        dialogControls.fontNameListBox.onChange = function () {
            if (dialogControls.fontNameListBox.selection !== null) {
                selectTextFramesByFontMatch(editSession.doc, dialogControls.fontNameListBox.selection.psNameKey);
            }
        };

        /**
         * 名前一覧タブを更新する
         * @returns {void}
         */
        function refreshInfoTabsHandler() { refreshInfoTabs(editSession); }
        dialogControls.cbFontPS.onClick = refreshInfoTabsHandler;
        dialogControls.cbFontFamily.onClick = refreshInfoTabsHandler;
        dialogControls.cbFontStyle.onClick = refreshInfoTabsHandler;

        /* 段落書式保持とプレビューの排他制御 / Make paragraph formatting and preview mutually exclusive */
        dialogControls.cbKeepFormat.onClick = function () {
            if (dialogControls.cbKeepFormat.value && dialogControls.cbPreview.value) {
                dialogControls.cbPreview.value = false;
                updatePreview(editSession);
            }
            if (!isPreviewUnsupported()) {
                dialogControls.cbPreview.enabled = !dialogControls.cbKeepFormat.value;
            }
        };

        /* Shift+Enter でソフト改行文字を挿入 / Insert the soft-break placeholder with Shift+Enter */
        dialogControls.textEditBox.addEventListener('keydown', function (keyEvent) {
            var isShift = ScriptUI.environment.keyboardState.shiftKey;
            if (isShift && keyEvent.keyName === 'Enter') {
                this.textselection = SOFT_BREAK;
                keyEvent.preventDefault();
            }
        });

        dialogControls.textEditBox.onChanging = function () { updatePreview(editSession); };
        dialogControls.cbPreview.onClick = function () { updatePreview(editSession); };
        dialogControls.btnExportText.onClick = function () { exportTextToDesktop(editSession); };

        /**
         * テキスト一覧を更新する
         * @returns {void}
         */
        function updateListHandler() { updateList(editSession); }

        /**
         * テキスト一覧とシンボル内テキスト一覧を集め直す
         * @returns {void}
         */
        function updateAllListsHandler() {
            updateList(editSession);
            refreshSymbolList(editSession, true);
        }

        /* ラジオボタン・チェックボックス切り替え時に更新 / Refresh when radio buttons or checkboxes change */
        dialogControls.rbCurrentArtboard.onClick = function () {
            dialogControls.cbWholeDocument.enabled = false;
            dialogControls.cbWholeDocument.value = false;
            updateAllListsHandler();
        };
        dialogControls.rbAllArtboards.onClick = function () {
            dialogControls.cbWholeDocument.enabled = true;
            updateAllListsHandler();
        };
        dialogControls.cbMergeDuplicates.onClick = updateListHandler;

        dialogControls.cbWholeDocument.onClick = updateAllListsHandler;
        dialogControls.cbIncludeCommentLayers.onClick = updateAllListsHandler;
        dialogControls.cbIncludeLocked.onClick = updateAllListsHandler;
        dialogControls.cbIncludeHidden.onClick = updateAllListsHandler;
        dialogControls.rbSortNone.onClick = updateListHandler;
        dialogControls.rbSortPosition.onClick = updateListHandler;
        dialogControls.rbSortAlphabetical.onClick = updateListHandler;
        dialogControls.rbLayerScopeAll.onClick = refreshInfoTabsHandler;
        dialogControls.rbLayerScopeTop.onClick = refreshInfoTabsHandler;
        dialogControls.rbArtboardScopeNumbered.onClick = refreshInfoTabsHandler;
        dialogControls.rbArtboardScopeRaw.onClick = refreshInfoTabsHandler;

        /* キャンセルボタンで閉じる / Close the dialog when Cancel is pressed */
        dialogControls.btnCancel.onClick = function () {
            dialogControls.textScopeDialog.close();
        };

        /* OKボタンで現在の編集を反映して閉じる / Apply the current edit and close when OK is pressed */
        dialogControls.btnOK.onClick = function () {
            if (editSession.isPreviewApplied && dialogControls.cbPreview.value) {
                app.undo();
                editSession.isPreviewApplied = false;
            }
            applyCurrentEdit(editSession);
            dialogControls.textScopeDialog.close();
        };

        /* ダイアログを閉じるとき、プレビュー中ならundoで元に戻す / Undo the preview when the dialog closes */
        dialogControls.textScopeDialog.onClose = function () {
            try {
                if (editSession.isPreviewApplied) app.undo();
                editSession.isPreviewApplied = false;
            } catch (err) { }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを開き、テキストの一覧と編集を扱う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        /* 編集の状態 / Edit session state */
        var editSession = {
            doc: app.activeDocument,
            dialogControls: null,
            textFrameList: [],        /* 一覧に並ぶテキストフレーム / Text frames listed */
            duplicateMap: [],         /* 行ごとの同じ内容の全フレーム / All frames with the same contents per row */
            symbolTextCache: {},      /* 範囲と条件ごとのシンボル内テキスト / Symbol texts per scope and options */
            isPreviewApplied: false   /* プレビューを1段適用中か / Whether one preview step is applied */
        };
        try {
            editSession.dialogControls = buildDialog();
            bindDialogEvents(editSession);

            /* 初回収集（updateList が名前一覧タブも更新する）/ Initial collection (updateList also refreshes the name tabs) */
            updateList(editSession);
            refreshSymbolList(editSession, false);

            editSession.dialogControls.textScopeDialog.show();
        } finally {
            if (editSession.isPreviewApplied) {
                try {
                    app.undo();
                } catch (err) { }
                editSession.isPreviewApplied = false;
            }
        }
    }

    main();

})();
