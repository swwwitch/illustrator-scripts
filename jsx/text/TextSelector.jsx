#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストフレームを、複数の条件で一括選択します。
属性・テキストの種類・文字列で絞り込み、選択後に非表示やレイヤー移動、一括編集も行えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md

### Overview

Selects text frames across the document by a combination of conditions.
You can filter by attribute, text kind and string, then hide, move to a layer, or bulk-edit what was selected.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextSelector";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「選択後に移動」先のレイヤー名 / Destination layer name for "move after selection" */
    var TEXT_LAYER_NAME = "_text";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];          /* パネルの余白 / panel margins */
    var OUTER_PANEL_MARGINS = [15, 20, 15, 15];    /* 選択条件・対象アートボードパネルの余白 / margins of the outer panels */
    var ATTRIBUTE_LABEL_WIDTH = 150;               /* 属性ラジオの幅 / width of the attribute radios */
    var PREVIEW_CHARACTERS = 24;                   /* 属性プレビューの文字数 / width of an attribute preview */
    var KEYWORD_CHARACTERS = 22;                   /* 検索文字列欄の文字数 / width of the keyword field */
    var BULK_INPUT_SIZE = [300, 70];               /* 一括編集の入力欄 / size of the bulk edit field */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale || "").toLowerCase().indexOf("ja") === 0 ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "テキストを選択", en: "Select Text" },
            bulkEditTitle: { ja: "テキストを一括変更", en: "Bulk Edit Text" }
        },
        panel: {
            selection: { ja: "選択条件", en: "Selection Criteria" },
            artboardScope: { ja: "対象とするアートボード", en: "Target Artboard" },
            attribute: { ja: "属性で選択", en: "Select by Attribute" },
            textType: { ja: "テキストの種類", en: "Text Type" },
            textMatch: { ja: "文字列で選択", en: "Select by String" },
            postProcess: { ja: "選択後の処理", en: "After Selection" }
        },
        radio: {
            artboardAll: { ja: "すべて", en: "All" },
            artboardCurrent: { ja: "現在のアートボードのみ", en: "Current Artboard Only" },
            fontFamily: { ja: "フォントファミリー", en: "Font Family" },
            fontFamilyStyle: { ja: "+ スタイル", en: "+ Style" },
            fontFamilyStyleSize: { ja: "+ スタイルとサイズ", en: "+ Style and Size" },
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            textFillColor: { ja: "テキストカラー", en: "Text Color" },
            opacity: { ja: "不透明度", en: "Opacity" },
            allText: { ja: "すべて", en: "All" },
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            pathText: { ja: "パス上文字", en: "Path Text" },
            exactMatch: { ja: "完全一致", en: "Exact" },
            containsMatch: { ja: "部分一致", en: "Contains" },
            startsWith: { ja: "先頭一致", en: "Starts With" },
            endsWith: { ja: "末尾一致", en: "Ends With" },
            regexMatch: { ja: "正規表現", en: "Regex" },
            noPostProcess: { ja: "なし", en: "None" },
            hideAfterSelection: { ja: "選択後に非表示", en: "Hide after selection" },
            hideOthers: { ja: "選択したテキスト以外を非表示", en: "Hide all but selected" },
            moveToTextLayer: { ja: "選択後に「_text」レイヤーへ移動", en: "Move to “_text” layer after selection" },
            bulkEdit: { ja: "一括編集", en: "Bulk edit" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            fontFamily: {
                ja: "選択中のテキストと同じフォントファミリーのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family as the selected text using Illustrator’s built-in command."
            },
            fontFamilyStyle: {
                ja: "選択中のテキストと同じフォントファミリー＋スタイルのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family and style as the selected text using Illustrator’s built-in command."
            },
            fontFamilyStyleSize: {
                ja: "選択中のテキストと同じフォントファミリー＋スタイル＋サイズのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family, style, and size as the selected text using Illustrator’s built-in command."
            },
            fontSize: {
                ja: "選択中のテキストと同じフォントサイズのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font size as the selected text using Illustrator’s built-in command."
            },
            textFillColor: {
                ja: "選択中のテキストと同じテキストカラーのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same text color as the selected text using Illustrator’s built-in command."
            },
            opacity: {
                ja: "選択中のオブジェクトと同じ不透明度のオブジェクトを、Illustrator標準コマンドで検索します。テキスト以外のオブジェクトも対象になります。",
                en: "Find objects with the same opacity as the selected object using Illustrator’s built-in command. Non-text objects are also matched."
            },
            allText: {
                ja: "ドキュメント内のすべてのテキストフレームを選択します。ショートカット：Option + Q",
                en: "Select all text frames in the document. Shortcut: Option + Q"
            },
            pointText: { ja: "ポイント文字だけを選択します。ショートカット：Option + W", en: "Select point text only. Shortcut: Option + W" },
            areaText: { ja: "エリア内文字だけを選択します。ショートカット：Option + E", en: "Select area text only. Shortcut: Option + E" },
            pathText: { ja: "パス上文字だけを選択します。", en: "Select path text only." },
            keywordInput: { ja: "検索対象にする文字列を入力します。空欄では実行できません。", en: "Enter the search string. This cannot be empty." },
            exactMatch: {
                ja: "入力した文字列と完全に一致するテキストを選択します。大小区別あり。ショートカット：Option + A",
                en: "Select text that exactly matches the entered string. Case-sensitive. Shortcut: Option + A"
            },
            startsWith: {
                ja: "入力した文字列で始まるテキストを選択します。大小区別あり。ショートカット：Option + B",
                en: "Select text that starts with the entered string. Case-sensitive. Shortcut: Option + B"
            },
            endsWith: {
                ja: "入力した文字列で終わるテキストを選択します。大小区別あり。ショートカット：Option + D",
                en: "Select text that ends with the entered string. Case-sensitive. Shortcut: Option + D"
            },
            containsMatch: {
                ja: "入力した文字列を含むテキストを選択します。大小区別あり。ショートカット：Option + I",
                en: "Select text that contains the entered string. Case-sensitive. Shortcut: Option + I"
            },
            regexMatch: {
                ja: "入力した正規表現に一致するテキストを選択します。大小区別あり。ショートカット：Option + R",
                en: "Select text that matches the entered regular expression. Case-sensitive. Shortcut: Option + R"
            },
            artboardAll: { ja: "ドキュメント内のすべてのアートボードを対象に選択します。", en: "Select across all artboards in the document." },
            artboardCurrent: {
                ja: "現在のアートボードに重なるオブジェクトだけを対象に選択します。",
                en: "Select only objects that overlap the active artboard."
            },
            noPostProcess: { ja: "選択後に追加処理を行いません。", en: "Do not apply any additional processing after selection." },
            hideAfterSelection: { ja: "選択されたテキストを非表示にします。", en: "Hide the selected text." },
            hideOthers: {
                ja: "選択されたテキスト以外のオブジェクトを非表示にします。ロックされたオブジェクトは除きます。",
                en: "Hide all objects except the selected text. Locked objects are left untouched."
            },
            moveToTextLayer: {
                ja: "選択されたテキストを「_text」レイヤーへ移動します。レイヤーがない場合は作成します。",
                en: "Move the selected text to the “_text” layer. The layer will be created if it does not exist."
            },
            bulkEdit: {
                ja: "選択されたテキストの内容をまとめて置き換えます。書式は維持されます。",
                en: "Replace contents of all selected text frames at once. Formatting is preserved."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
            emptyKeyword: { ja: "検索文字列を入力してください。", en: "Enter a search string." },
            invalidRegex: { ja: "正規表現が正しくありません。", en: "The regular expression is invalid." },
            noMatch: { ja: "条件に一致するテキストが見つかりませんでした。", en: "No text matched the criteria." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
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
    // 選択中のテキストの情報 / Current text selection
    // =========================================

    /**
     * 選択（配列・単体・TextRange）を配列にそろえる
     * @param {Object} docSelection - doc.selection の値
     * @returns {Array} 選択中のオブジェクトの配列
     */
    function collectSelectionAsArray(docSelection) {
        var selectedItems = [];
        if (!docSelection) {
            return selectedItems;
        }

        if (docSelection.typename) {
            selectedItems.push(docSelection);
            return selectedItems;
        }

        if (typeof docSelection.length === "number") {
            for (var i = 0; i < docSelection.length; i++) {
                selectedItems.push(docSelection[i]);
            }
            return selectedItems;
        }

        selectedItems.push(docSelection);
        return selectedItems;
    }

    /**
     * 選択中のオブジェクトからテキスト範囲を取り出す
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {TextRange|null} テキスト範囲（無ければ null）
     */
    function getTextRangeFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return null;
        }

        /* テキスト以外ではプロパティの参照で例外になることがある / non-text items may throw on these properties */
        try {
            if (selectionItem.typename === "TextFrame") {
                return selectionItem.textRange;
            }
            if (selectionItem.characterAttributes) {
                return selectionItem;
            }
            if (selectionItem.textRange) {
                return selectionItem.textRange;
            }
        } catch (textRangeError) {
            return null;
        }

        return null;
    }

    /**
     * 選択中のオブジェクトの文字列を取り出す
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {string} 文字列（無ければ空文字）
     */
    function getTextContentsFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return "";
        }

        /* contents を持たないオブジェクトでは例外になることがある / may throw on items without contents */
        try {
            if (typeof selectionItem.contents === "string") {
                return selectionItem.contents;
            }
        } catch (contentsError) {
        }

        var textRange = getTextRangeFromSelectionItem(selectionItem);
        if (textRange && typeof textRange.contents === "string") {
            return textRange.contents;
        }

        return "";
    }

    /**
     * 選択中の最初のテキストの文字列を返す（検索文字列の初期値）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {string} 文字列（無ければ空文字）
     */
    function getSelectedTextString(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var textContents = getTextContentsFromSelectionItem(selectedItems[i]);
            if (textContents) {
                return textContents;
            }
        }
        return "";
    }

    /**
     * 属性プレビューの初期値（未取得を表す「—」）を返す
     * @returns {Object} family / familyStyle / familyStyleSize / size / opacity
     */
    function makeEmptyAttributePreview() {
        return {
            family: "—",
            familyStyle: "—",
            familyStyleSize: "—",
            size: "—",
            opacity: "—"
        };
    }

    /**
     * テキスト範囲から文字属性のプレビューを作る
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {Object} makeEmptyAttributePreview() と同じ形のプレビュー
     */
    function getTextAttributePreviewFromTextRange(textRange) {
        var attributePreview = makeEmptyAttributePreview();

        /* フォントが見つからないなどで属性が読めないことがある / attributes may be unreadable, e.g. a missing font */
        try {
            var characterAttributes = textRange.characterAttributes;
            var textFont = characterAttributes.textFont;
            var fontSize = characterAttributes.size;
            var familyName = textFont.family || "";
            var styleName = textFont.style || "";
            var sizeText = fontSize ? String(Math.round(fontSize * 100) / 100) + " pt" : "—";

            attributePreview.family = familyName || "—";
            attributePreview.familyStyle = familyName ? familyName + (styleName ? " / " + styleName : "") : "—";
            attributePreview.familyStyleSize = attributePreview.familyStyle !== "—" ? attributePreview.familyStyle + " / " + sizeText : "—";
            attributePreview.size = sizeText;
        } catch (attributePreviewError) {
            return attributePreview;
        }

        return attributePreview;
    }

    /**
     * 選択中のオブジェクトから不透明度のプレビューを作る
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {string} 不透明度の表示（無ければ「—」）
     */
    function getOpacityPreviewFromSelectionItem(selectionItem) {
        /* TextRange などでは opacity の参照で例外になることがある / may throw on a TextRange and the like */
        try {
            if (selectionItem && typeof selectionItem.opacity === "number") {
                return String(Math.round(selectionItem.opacity * 100) / 100) + " %";
            }
        } catch (opacityPreviewError) {
        }
        return "—";
    }

    /**
     * 選択中の最初のテキストから文字属性と不透明度のプレビューを作る
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {Object} makeEmptyAttributePreview() と同じ形のプレビュー
     */
    function getSelectedTextAttributePreview(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var textRange = getTextRangeFromSelectionItem(selectedItems[i]);
            if (textRange) {
                var rangePreview = getTextAttributePreviewFromTextRange(textRange);
                rangePreview.opacity = getOpacityPreviewFromSelectionItem(selectedItems[i]);
                return rangePreview;
            }
        }

        return makeEmptyAttributePreview();
    }

    // =========================================
    // 条件による選択 / Selection by criteria
    // =========================================

    /**
     * 文字列が検索文字列に一致するかを返す
     * @param {string} frameText - テキストフレームの文字列
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - "exact" / "startsWith" / "endsWith" / "contains" / "regex"
     * @param {RegExp|null} keywordPattern - 正規表現モードのときの正規表現
     * @returns {boolean} 一致すれば true
     */
    function textMatches(frameText, keyword, matchMode, keywordPattern) {
        if (matchMode === "regex") {
            return keywordPattern ? keywordPattern.test(frameText) : false;
        }
        if (matchMode === "exact") {
            return frameText === keyword;
        }
        if (matchMode === "startsWith") {
            return frameText.indexOf(keyword) === 0;
        }
        if (matchMode === "endsWith") {
            return frameText.lastIndexOf(keyword) === frameText.length - keyword.length;
        }
        if (matchMode === "contains") {
            return frameText.indexOf(keyword) !== -1;
        }
        return false;
    }

    /**
     * 文字列条件の入力を確かめる（空欄・不正な正規表現はメッセージを出す）
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - 一致のモード
     * @returns {Object|null} { regex: RegExp|null }。入力が不正なら null
     */
    function validateTextMatchInput(keyword, matchMode) {
        if (!keyword) {
            alert(getLabel("alert.emptyKeyword"));
            return null;
        }
        if (matchMode === "regex") {
            /* 不正なパターンは RegExp が例外を投げる / RegExp throws on an invalid pattern */
            try {
                return { regex: new RegExp(keyword) };
            } catch (regexError) {
                alert(getLabel("alert.invalidRegex"));
                return null;
            }
        }
        return { regex: null };
    }

    /**
     * テキストの種類に対応する判定関数を作る
     * @param {string} textType - "all" / "point" / "area" / "path"
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildTextTypePredicate(textType) {
        if (textType === "point") {
            return function (textFrame) { return textFrame.kind === TextType.POINTTEXT; };
        }
        if (textType === "area") {
            return function (textFrame) { return textFrame.kind === TextType.AREATEXT; };
        }
        if (textType === "path") {
            return function (textFrame) { return textFrame.kind === TextType.PATHTEXT; };
        }
        /* "all" は全件一致 / "all" matches everything */
        return function () { return true; };
    }

    /**
     * 現在のアートボードの矩形を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]|null} [左, 上, 右, 下]（取れなければ null）
     */
    function getActiveArtboardRect(doc) {
        /* アートボードの取得に失敗したら絞り込まない / skip the filter when the artboard cannot be read */
        try {
            var activeIndex = doc.artboards.getActiveArtboardIndex();
            return doc.artboards[activeIndex].artboardRect; /* [left, top, right, bottom] */
        } catch (artboardRectError) {
            return null;
        }
    }

    /**
     * オブジェクトが矩形に重なるかを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @param {number[]|null} artboardRect - [左, 上, 右, 下]（null なら常に true）
     * @returns {boolean} 重なれば true
     */
    function isItemWithinArtboardRect(pageItem, artboardRect) {
        if (!artboardRect) {
            return true;
        }
        /* 境界が読めないオブジェクトは残す / keep items whose bounds cannot be read */
        try {
            var itemBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
            var overlapsHorizontally = itemBounds[2] >= artboardRect[0] && itemBounds[0] <= artboardRect[2];
            var overlapsVertically = itemBounds[1] >= artboardRect[3] && itemBounds[3] <= artboardRect[1];
            return overlapsHorizontally && overlapsVertically;
        } catch (boundsError) {
            return true;
        }
    }

    /**
     * 現在のアートボードに重なるオブジェクトだけに絞り込む
     * @param {Array} pageItems - 対象のオブジェクト
     * @returns {Array} 絞り込んだオブジェクト
     */
    function filterItemsByActiveArtboard(pageItems) {
        var artboardRect = getActiveArtboardRect(app.activeDocument);
        var itemsOnArtboard = [];
        for (var i = 0; i < pageItems.length; i++) {
            if (isItemWithinArtboardRect(pageItems[i], artboardRect)) {
                itemsOnArtboard.push(pageItems[i]);
            }
        }
        return itemsOnArtboard;
    }

    /**
     * 判定関数に一致するテキストフレームを選択する
     * @param {Function} predicate - テキストフレームを受け取って true/false を返す関数
     * @param {string} artboardScope - "all" / "current"
     * @returns {number} 選択した件数
     */
    function selectTextFrames(predicate, artboardScope) {
        var doc = app.activeDocument;
        var allTextFrames = doc.textFrames;
        var matchedFrames = [];

        for (var i = 0; i < allTextFrames.length; i++) {
            if (predicate(allTextFrames[i])) {
                matchedFrames.push(allTextFrames[i]);
            }
        }

        if (artboardScope === "current") {
            matchedFrames = filterItemsByActiveArtboard(matchedFrames);
        }

        doc.selection = matchedFrames;
        return matchedFrames.length;
    }

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 一括編集の文字列を入力するダイアログを表示する
     * @returns {string|null} 入力した文字列（キャンセル時は null）
     */
    function showBulkEditDialog() {
        var bulkDialog = new Window("dialog", getLabel("dialog.bulkEditTitle"));
        bulkDialog.orientation = "column";
        bulkDialog.alignChildren = "fill";
        bulkDialog.spacing = 10;
        bulkDialog.margins = 16;

        var replacementInput = bulkDialog.add("edittext", undefined, "", {
            multiline: true,
            scrolling: true
        });
        replacementInput.preferredSize = BULK_INPUT_SIZE;
        replacementInput.active = true;

        var bulkBtnRowGroup = bulkDialog.add("group");
        bulkBtnRowGroup.alignment = "right";
        var btnBulkCancel = bulkBtnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnBulkOK = bulkBtnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnBulkOK.onClick = function () {
            bulkDialog.close(1);
        };
        btnBulkCancel.onClick = function () {
            bulkDialog.close(0);
        };

        if (bulkDialog.show() !== 1) {
            return null;
        }
        return replacementInput.text;
    }

    /**
     * テキストフレームの内容をまとめて置き換える（文字ごとの書式は元の文字数ぶん引き継ぐ）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function bulkEditTextFrames(textFrames) {
        if (textFrames.length === 0) {
            return;
        }

        var replacementText = showBulkEditDialog();
        if (replacementText === null) {
            return;
        }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var originalRange = textFrame.textRange;
            var originalLength = originalRange.length;
            if (originalLength === 0) {
                continue;
            }

            textFrame.contents = replacementText;

            var newRange = textFrame.textRange;
            var sharedLength = Math.min(originalLength, newRange.length);
            for (var j = 0; j < sharedLength; j++) {
                /* 書式を写せない文字は飛ばす / skip characters whose formatting cannot be copied */
                try {
                    newRange.characters[j].characterAttributes = originalRange.characters[j].characterAttributes;
                } catch (attrCopyError) {
                }
            }
        }

        app.redraw();
    }

    /**
     * レイヤーを名前で取得し、無ければ作る。ロック解除・表示にして、元の状態も返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Object} { layer, originalLocked, originalVisible }
     */
    function getOrCreateLayerByName(doc, layerName) {
        var targetLayer;
        /* getByName は見つからないと例外 / getByName throws when the layer is missing */
        try {
            targetLayer = doc.layers.getByName(layerName);
        } catch (layerFindError) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }

        var originalLocked = targetLayer.locked;
        var originalVisible = targetLayer.visible;

        targetLayer.locked = false;
        targetLayer.visible = true;

        return {
            layer: targetLayer,
            originalLocked: originalLocked,
            originalVisible: originalVisible
        };
    }

    /**
     * 選択中の子孫を持つかを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 子孫に選択中のものがあれば true
     */
    function containsSelectedDescendant(pageItem) {
        var childItems;
        /* pageItems を持たないオブジェクトでは例外になる / items without pageItems throw */
        try {
            childItems = pageItem.pageItems;
        } catch (childAccessError) {
            return false;
        }
        if (!childItems || childItems.length === 0) {
            return false;
        }
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            /* selected を読めないオブジェクトは飛ばす / skip items whose selected state cannot be read */
            try {
                if (childItem.selected) {
                    return true;
                }
            } catch (selectedReadError) {
            }
            if (containsSelectedDescendant(childItem)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 選択中以外のオブジェクトを非表示にする（ロック中と、選択を含むグループは残す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function hideItemsExceptSelection(doc) {
        var allPageItems = doc.pageItems;
        for (var i = 0; i < allPageItems.length; i++) {
            var pageItem = allPageItems[i];
            /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
            try {
                /* 選択中・ロック中・選択を内包するグループは残す / Keep selected, locked, or selection-containing items */
                if (pageItem.selected || pageItem.locked) {
                    continue;
                }
                if (containsSelectedDescendant(pageItem)) {
                    continue;
                }
                pageItem.hidden = true;
            } catch (hideOtherError) {
            }
        }
    }

    /**
     * 選択中のオブジェクトを非表示にする
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function hideSelectedItems(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
            try {
                selectedItems[i].hidden = true;
            } catch (hideItemError) {
            }
        }
    }

    /**
     * 選択中のオブジェクトを「_text」レイヤーへ移動する（レイヤーのロック・表示状態は元に戻す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function moveItemsToTextLayer(doc, selectedItems) {
        var layerInfo = getOrCreateLayerByName(doc, TEXT_LAYER_NAME);
        var textLayer = layerInfo.layer;
        for (var j = 0; j < selectedItems.length; j++) {
            /* ロック状態や親レイヤーの状態によって移動できない場合はスキップ / Skip items that cannot be moved because of lock state or parent layer state */
            try {
                selectedItems[j].locked = false;
                selectedItems[j].move(textLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (moveItemError) {
            }
        }
        /* 元のロック・可視状態へ戻す（新規作成時はデフォルト値の戻りで実質 no-op） / Restore original lock/visibility (no-op when newly created) */
        try {
            textLayer.locked = layerInfo.originalLocked;
            textLayer.visible = layerInfo.originalVisible;
        } catch (restoreLayerError) {
        }
    }

    /**
     * 選択後の処理を行う
     * @param {string} postProcessMode - "" / "hide" / "hideOthers" / "moveToTextLayer" / "bulkEdit"
     * @returns {void}
     */
    function applyPostProcessToSelection(postProcessMode) {
        if (!postProcessMode) {
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = collectSelectionAsArray(doc.selection);
        if (selectedItems.length === 0) {
            return;
        }

        if (postProcessMode === "hide") {
            hideSelectedItems(selectedItems);
        } else if (postProcessMode === "hideOthers") {
            hideItemsExceptSelection(doc);
        } else if (postProcessMode === "bulkEdit") {
            var textFramesOnly = [];
            for (var k = 0; k < selectedItems.length; k++) {
                if (selectedItems[k].typename === "TextFrame") {
                    textFramesOnly.push(selectedItems[k]);
                }
            }
            bulkEditTextFrames(textFramesOnly);
        } else if (postProcessMode === "moveToTextLayer") {
            moveItemsToTextLayer(doc, selectedItems);
        }
    }

    /**
     * 選択件数を確かめてから後処理を行う（0件ならメッセージ）
     * @param {number} selectedCount - 選択した件数
     * @param {string} postProcessMode - 後処理のモード
     * @returns {void}
     */
    function finalizeSelection(selectedCount, postProcessMode) {
        if (selectedCount === 0) {
            alert(getLabel("alert.noMatch"));
            return;
        }
        applyPostProcessToSelection(postProcessMode);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルに共通のレイアウトを設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupPanelLayout(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    /**
     * ヘルプチップを設定する（文言が空なら何もしない）
     * @param {Object} targetControl - 対象のコントロール
     * @param {string} tipText - ヘルプチップの文言
     * @returns {void}
     */
    function setHelpTip(targetControl, tipText) {
        if (targetControl && tipText) {
            targetControl.helpTip = tipText;
        }
    }

    /**
     * ラジオボタンの行を追加する。previewText を渡すと右に値のプレビューを添える
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelText - ラジオボタンのラベル
     * @param {string|null} previewText - プレビューの文字列（null なら添えない）
     * @param {number} [labelWidth] - ラジオボタンの幅
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadioRow(parentPanel, labelText, previewText, labelWidth) {
        var rowGroup = parentPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = "fill";

        var radioButton = rowGroup.add("radiobutton", undefined, labelText);
        if (typeof labelWidth === "number") {
            radioButton.preferredSize.width = labelWidth;
        }
        if (previewText !== null) {
            var previewLabel = rowGroup.add("statictext", undefined, previewText);
            previewLabel.alignment = ["fill", "center"];
            previewLabel.characters = PREVIEW_CHARACTERS;
        }
        return radioButton;
    }

    /**
     * 縦並び・左揃えの列グループを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "left";
        return columnGroup;
    }

    /**
     * ラジオボタンを、親をまたいで排他にする
     * @param {RadioButton[]} radioButtons - 排他にするラジオボタン
     * @returns {void}
     */
    function setupExclusiveRadioButtons(radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            (function (currentRadioButton) {
                currentRadioButton.onClick = function () {
                    for (var j = 0; j < radioButtons.length; j++) {
                        if (radioButtons[j] !== currentRadioButton) {
                            radioButtons[j].value = false;
                        }
                    }
                };
            })(radioButtons[i]);
        }
    }

    /**
     * 指定したラジオボタンだけを選択する
     * @param {RadioButton[]} radioButtons - 同じ組のラジオボタン
     * @param {RadioButton} targetRadioButton - 選択するラジオボタン
     * @returns {void}
     */
    function selectExclusiveRadioButton(radioButtons, targetRadioButton) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (radioButtons[i] === targetRadioButton);
        }
    }

    /**
     * Option＋キーで選択条件を切り替えるハンドラーを付ける（検索文字列の入力中は無効）
     * @param {Object} ui - buildDialog() の結果
     * @returns {void}
     */
    function addSelectionKeyHandler(ui) {
        var shortcutRadios = {
            Q: ui.rbAllText,
            W: ui.rbPointText,
            E: ui.rbAreaText,
            A: ui.rbExactMatch,
            B: ui.rbStartsWith,
            D: ui.rbEndsWith,
            I: ui.rbContainsMatch,
            R: ui.rbRegexMatch
        };
        ui.window.addEventListener("keydown", function (event) {
            if (ui.keywordInput && ui.keywordInput.active) {
                return;
            }
            if (!event.altKey) {
                return;
            }

            var targetRadioButton = shortcutRadios.hasOwnProperty(event.keyName) ? shortcutRadios[event.keyName] : null;
            if (targetRadioButton) {
                selectExclusiveRadioButton(ui.selectionRadios, targetRadioButton);
                event.preventDefault();
            }
        });
    }

    /**
     * 選択条件パネル（対象アートボード・属性・テキストの種類・文字列）を追加する
     * @param {Object} ui - コントロールの格納先
     * @param {Object} initialState - 開いた時点の選択の情報（hasSelection / keyword / attributePreview）
     * @returns {void}
     */
    function addSelectionPanel(ui, initialState) {
        var selectionPanel = ui.window.add("panel", undefined, getLabel("panel.selection"));
        setupPanelLayout(selectionPanel, 10);
        selectionPanel.margins = OUTER_PANEL_MARGINS;

        /* 対象アートボードパネル / Target artboard panel */
        var artboardScopePanel = selectionPanel.add("panel", undefined, getLabel("panel.artboardScope"));
        artboardScopePanel.orientation = "row";
        artboardScopePanel.alignChildren = ["left", "center"];
        artboardScopePanel.alignment = "fill";
        artboardScopePanel.margins = OUTER_PANEL_MARGINS;
        artboardScopePanel.spacing = 20;

        ui.rbArtboardAll = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardAll"));
        ui.rbArtboardCurrent = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardCurrent"));
        ui.rbArtboardAll.value = true;
        setHelpTip(ui.rbArtboardAll, getLabel("tooltip.artboardAll"));
        setHelpTip(ui.rbArtboardCurrent, getLabel("tooltip.artboardCurrent"));
        setupExclusiveRadioButtons([ui.rbArtboardAll, ui.rbArtboardCurrent]);

        /* 属性選択パネル / Attribute selection panel */
        var attributePanel = selectionPanel.add("panel", undefined, getLabel("panel.attribute"));
        setupPanelLayout(attributePanel, 6);
        attributePanel.enabled = initialState.hasSelection;

        var attributePreview = initialState.attributePreview;
        ui.rbFontFamily = addRadioRow(attributePanel, getLabel("radio.fontFamily"), attributePreview.family, ATTRIBUTE_LABEL_WIDTH);
        ui.rbFontFamilyStyle = addRadioRow(attributePanel, getLabel("radio.fontFamilyStyle"), attributePreview.familyStyle, ATTRIBUTE_LABEL_WIDTH);
        ui.rbFontFamilyStyleSize = addRadioRow(attributePanel, getLabel("radio.fontFamilyStyleSize"), attributePreview.familyStyleSize, ATTRIBUTE_LABEL_WIDTH);
        ui.rbFontSize = addRadioRow(attributePanel, getLabel("radio.fontSize"), attributePreview.size, ATTRIBUTE_LABEL_WIDTH);
        ui.rbTextFillColor = addRadioRow(attributePanel, getLabel("radio.textFillColor"), null, ATTRIBUTE_LABEL_WIDTH);
        ui.rbOpacity = addRadioRow(attributePanel, getLabel("radio.opacity"), attributePreview.opacity, ATTRIBUTE_LABEL_WIDTH);

        setHelpTip(ui.rbFontFamily, getLabel("tooltip.fontFamily"));
        setHelpTip(ui.rbFontFamilyStyle, getLabel("tooltip.fontFamilyStyle"));
        setHelpTip(ui.rbFontFamilyStyleSize, getLabel("tooltip.fontFamilyStyleSize"));
        setHelpTip(ui.rbFontSize, getLabel("tooltip.fontSize"));
        setHelpTip(ui.rbTextFillColor, getLabel("tooltip.textFillColor"));
        setHelpTip(ui.rbOpacity, getLabel("tooltip.opacity"));

        /* テキスト種類と文字列条件を横並びに配置 / Arrange text type and string condition panels side by side */
        var textConditionGroup = selectionPanel.add("group");
        textConditionGroup.orientation = "row";
        textConditionGroup.alignChildren = "fill";
        textConditionGroup.alignment = "fill";

        /* テキスト種類パネル / Text type panel */
        var textTypePanel = textConditionGroup.add("panel", undefined, getLabel("panel.textType"));
        setupPanelLayout(textTypePanel, 6);

        ui.rbAllText = textTypePanel.add("radiobutton", undefined, getLabel("radio.allText"));
        ui.rbPointText = textTypePanel.add("radiobutton", undefined, getLabel("radio.pointText"));
        ui.rbAreaText = textTypePanel.add("radiobutton", undefined, getLabel("radio.areaText"));
        ui.rbPathText = textTypePanel.add("radiobutton", undefined, getLabel("radio.pathText"));

        setHelpTip(ui.rbAllText, getLabel("tooltip.allText"));
        setHelpTip(ui.rbPointText, getLabel("tooltip.pointText"));
        setHelpTip(ui.rbAreaText, getLabel("tooltip.areaText"));
        setHelpTip(ui.rbPathText, getLabel("tooltip.pathText"));

        /* 初期選択を設定（選択があれば「＋スタイルとサイズ」、なければ「すべて」） / Set initial selection */
        if (initialState.hasSelection) {
            ui.rbFontFamilyStyleSize.value = true;
        } else {
            ui.rbAllText.value = true;
        }

        /* 文字列条件パネル / String condition panel */
        var textMatchPanel = textConditionGroup.add("panel", undefined, getLabel("panel.textMatch"));
        setupPanelLayout(textMatchPanel, 6);

        ui.keywordInput = textMatchPanel.add("edittext", undefined, initialState.keyword);
        ui.keywordInput.characters = KEYWORD_CHARACTERS;
        setHelpTip(ui.keywordInput, getLabel("tooltip.keywordInput"));

        var textMatchOptionsGroup = textMatchPanel.add("group");
        textMatchOptionsGroup.orientation = "row";
        textMatchOptionsGroup.alignChildren = "top";
        textMatchOptionsGroup.margins = [0, 10, 0, 0];
        textMatchOptionsGroup.spacing = 22;

        var textMatchLeftColumnGroup = addColumnGroup(textMatchOptionsGroup);
        var textMatchCenterColumnGroup = addColumnGroup(textMatchOptionsGroup);
        var textMatchRightColumnGroup = addColumnGroup(textMatchOptionsGroup);

        ui.rbExactMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, getLabel("radio.exactMatch"));
        ui.rbContainsMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, getLabel("radio.containsMatch"));
        ui.rbStartsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, getLabel("radio.startsWith"));
        ui.rbEndsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, getLabel("radio.endsWith"));
        ui.rbRegexMatch = textMatchRightColumnGroup.add("radiobutton", undefined, getLabel("radio.regexMatch"));

        setHelpTip(ui.rbExactMatch, getLabel("tooltip.exactMatch"));
        setHelpTip(ui.rbStartsWith, getLabel("tooltip.startsWith"));
        setHelpTip(ui.rbEndsWith, getLabel("tooltip.endsWith"));
        setHelpTip(ui.rbContainsMatch, getLabel("tooltip.containsMatch"));
        setHelpTip(ui.rbRegexMatch, getLabel("tooltip.regexMatch"));

        /* 選択条件のラジオは、パネルをまたいで排他にする / The criteria radios are exclusive across panels */
        ui.selectionRadios = [
            ui.rbFontFamily, ui.rbFontFamilyStyle, ui.rbFontFamilyStyleSize, ui.rbFontSize, ui.rbTextFillColor, ui.rbOpacity,
            ui.rbAllText, ui.rbPointText, ui.rbAreaText, ui.rbPathText,
            ui.rbExactMatch, ui.rbStartsWith, ui.rbEndsWith, ui.rbContainsMatch, ui.rbRegexMatch
        ];
        setupExclusiveRadioButtons(ui.selectionRadios);
    }

    /**
     * 選択後の処理パネルを追加する
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addPostProcessPanel(ui) {
        var postProcessPanel = ui.window.add("panel", undefined, getLabel("panel.postProcess"));
        setupPanelLayout(postProcessPanel, 6);
        /* ドキュメント内に TextFrame が 0 件なら後処理は無意味なのでディム / Disable post-process when document has no text frames */
        postProcessPanel.enabled = app.activeDocument.textFrames.length > 0;

        ui.rbNoPostProcess = postProcessPanel.add("radiobutton", undefined, getLabel("radio.noPostProcess"));
        ui.rbHide = postProcessPanel.add("radiobutton", undefined, getLabel("radio.hideAfterSelection"));
        ui.rbHideOthers = postProcessPanel.add("radiobutton", undefined, getLabel("radio.hideOthers"));
        ui.rbMove = postProcessPanel.add("radiobutton", undefined, getLabel("radio.moveToTextLayer"));
        ui.rbBulkEdit = postProcessPanel.add("radiobutton", undefined, getLabel("radio.bulkEdit"));

        ui.rbNoPostProcess.value = true;

        setHelpTip(ui.rbNoPostProcess, getLabel("tooltip.noPostProcess"));
        setHelpTip(ui.rbHide, getLabel("tooltip.hideAfterSelection"));
        setHelpTip(ui.rbHideOthers, getLabel("tooltip.hideOthers"));
        setHelpTip(ui.rbMove, getLabel("tooltip.moveToTextLayer"));
        setHelpTip(ui.rbBulkEdit, getLabel("tooltip.bulkEdit"));

        setupExclusiveRadioButtons([ui.rbNoPostProcess, ui.rbHide, ui.rbHideOthers, ui.rbMove, ui.rbBulkEdit]);
    }

    /**
     * ダイアログを組み立てる（OK/キャンセルの処理は main() で付ける）
     * @param {Object} initialState - 開いた時点の選択の情報（hasSelection / keyword / attributePreview）
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildDialog(initialState) {
        var ui = {};
        ui.window = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);

        addSelectionPanel(ui, initialState);
        addSelectionKeyHandler(ui);
        addPostProcessPanel(ui);

        /* ボタンエリア / Button area */
        var btnRowGroup = ui.window.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        ui.btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        ui.btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return ui;
    }

    /* 属性ラジオと、選択に使う Illustrator 標準コマンド / Attribute radios and the built-in commands they run */
    var ATTRIBUTE_COMMANDS = [
        { radioKey: "rbFontFamily", command: "Find Text Font Family menu item" },
        { radioKey: "rbFontFamilyStyle", command: "Find Text Font Family Style menu item" },
        { radioKey: "rbFontFamilyStyleSize", command: "Find Text Font Family Style Size menu item" },
        { radioKey: "rbFontSize", command: "Find Text Font Size menu item" },
        { radioKey: "rbTextFillColor", command: "Find Text Fill Color menu item" },
        { radioKey: "rbOpacity", command: "Find Opacity menu item" }
    ];

    /* 文字列の一致モードのラジオ（判定順） / String match radios, in checking order */
    var TEXT_MATCH_MODES = [
        { radioKey: "rbExactMatch", mode: "exact" },
        { radioKey: "rbStartsWith", mode: "startsWith" },
        { radioKey: "rbEndsWith", mode: "endsWith" },
        { radioKey: "rbContainsMatch", mode: "contains" },
        { radioKey: "rbRegexMatch", mode: "regex" }
    ];

    /* テキストの種類のラジオ（判定順） / Text type radios, in checking order */
    var TEXT_TYPES = [
        { radioKey: "rbAllText", textType: "all" },
        { radioKey: "rbPointText", textType: "point" },
        { radioKey: "rbAreaText", textType: "area" },
        { radioKey: "rbPathText", textType: "path" }
    ];

    /* 後処理のラジオ（判定順） / Post-process radios, in checking order */
    var POST_PROCESS_MODES = [
        { radioKey: "rbNoPostProcess", mode: "" },
        { radioKey: "rbHide", mode: "hide" },
        { radioKey: "rbHideOthers", mode: "hideOthers" },
        { radioKey: "rbMove", mode: "moveToTextLayer" },
        { radioKey: "rbBulkEdit", mode: "bulkEdit" }
    ];

    /**
     * ラジオの対応表から、オンになっている最初の行の値を返す
     * @param {Object} ui - buildDialog() の結果
     * @param {Object[]} radioTable - radioKey と値を持つ行の配列
     * @param {string} valueKey - 返す値のキー
     * @param {string} fallbackValue - どれもオフのときの値
     * @returns {string} 値
     */
    function readRadioChoice(ui, radioTable, valueKey, fallbackValue) {
        for (var i = 0; i < radioTable.length; i++) {
            if (ui[radioTable[i].radioKey].value) return radioTable[i][valueKey];
        }
        return fallbackValue;
    }

    /**
     * ダイアログの状態から設定を読み取る
     * @param {Object} ui - buildDialog() の結果
     * @returns {Object} postProcessMode / artboardScope / attributeCommand / textMatchMode / keyword / textType
     */
    function readDialogSettings(ui) {
        return {
            postProcessMode: readRadioChoice(ui, POST_PROCESS_MODES, "mode", ""),
            artboardScope: ui.rbArtboardCurrent.value ? "current" : "all",
            attributeCommand: readRadioChoice(ui, ATTRIBUTE_COMMANDS, "command", ""),
            textMatchMode: readRadioChoice(ui, TEXT_MATCH_MODES, "mode", ""),
            keyword: ui.keywordInput.text || "",
            textType: readRadioChoice(ui, TEXT_TYPES, "textType", "all")
        };
    }

    /**
     * 属性の標準コマンドで選択し、必要なら現在のアートボードに絞って後処理を行う
     * @param {Object} selectorSettings - readDialogSettings() の結果
     * @returns {void}
     */
    function selectByAttribute(selectorSettings) {
        app.executeMenuCommand(selectorSettings.attributeCommand);
        if (selectorSettings.artboardScope === "current") {
            var filteredAttributeSelection = filterItemsByActiveArtboard(collectSelectionAsArray(app.activeDocument.selection));
            app.activeDocument.selection = filteredAttributeSelection;
            finalizeSelection(filteredAttributeSelection.length, selectorSettings.postProcessMode);
            return;
        }
        applyPostProcessToSelection(selectorSettings.postProcessMode);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで条件を選び、テキストフレームを選択して後処理を行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var initialSelection = collectSelectionAsArray(app.activeDocument.selection);
        var ui = buildDialog({
            hasSelection: initialSelection.length > 0,
            keyword: getSelectedTextString(initialSelection),
            attributePreview: getSelectedTextAttributePreview(initialSelection)
        });

        /* OKボタン実行処理 / Handle OK button action */
        ui.btnOK.onClick = function () {
            var selectorSettings = readDialogSettings(ui);

            /* 属性で選択：Illustrator標準コマンドに選択を任せる / By attribute: let Illustrator's command perform the selection */
            if (selectorSettings.attributeCommand) {
                ui.window.close();
                selectByAttribute(selectorSettings);
                return;
            }

            /* 文字列で選択 / By string */
            if (selectorSettings.textMatchMode) {
                var keywordValidation = validateTextMatchInput(selectorSettings.keyword, selectorSettings.textMatchMode);
                if (!keywordValidation) {
                    return;
                }
                ui.window.close();
                var stringPredicate = function (textFrame) {
                    return textMatches(textFrame.contents || "", selectorSettings.keyword, selectorSettings.textMatchMode, keywordValidation.regex);
                };
                finalizeSelection(selectTextFrames(stringPredicate, selectorSettings.artboardScope), selectorSettings.postProcessMode);
                return;
            }

            /* テキストの種類で選択 / By text type */
            ui.window.close();
            var typePredicate = buildTextTypePredicate(selectorSettings.textType);
            finalizeSelection(selectTextFrames(typePredicate, selectorSettings.artboardScope), selectorSettings.postProcessMode);
        };

        /* キャンセルボタン処理 / Handle Cancel button action */
        ui.btnCancel.onClick = function () {
            ui.window.close();
        };

        ui.window.show();
    }

    main();

})();
