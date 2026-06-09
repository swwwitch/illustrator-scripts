#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview

Illustratorドキュメント内のテキストフレームを、複数の条件で一括選択するスクリプトです。

【選択条件】
- 属性で選択：選択中テキストを基準に、フォントファミリー／＋スタイル／＋サイズ／フォントサイズ／テキストカラー／不透明度をIllustrator標準コマンドで検索
- テキストの種類：すべて／ポイント文字／エリア内文字／パス上文字
- 文字列で選択：完全一致／部分一致／先頭一致／末尾一致／正規表現（大小区別あり）

【選択後の処理】
- なし
- 選択後に非表示
- 「_text」レイヤーへ移動（既存レイヤーのロック・可視状態は復元）
- 一括編集：書式（characterAttributes）を維持したまま、入力した文字列で内容を置換

【処理の流れ】
1. ダイアログで選択条件と選択後の処理を指定
2. 条件に一致するテキストフレームを選択（0件ならアラート）
3. 選択後の処理を適用（一括編集の場合はサブダイアログで入力）

Select text frames in an Illustrator document using multiple conditions.

[Selection criteria]
- By attribute: find text matching the selected text’s font family / + style / + size / font size / text color / opacity via Illustrator’s built-in commands
- By text type: All / Point Text / Area Text / Path Text
- By string: Exact / Contains / Starts With / Ends With / Regex (case-sensitive)

[After selection]
- None
- Hide
- Move to the “_text” layer (the original lock/visibility of an existing layer is restored)
- Bulk edit: replace contents while preserving formatting (characterAttributes)

[Flow]
1. Choose selection criteria and post-process in the dialog
2. Select matching text frames (alert if none matched)
3. Apply the post-process (bulk edit opens a sub-dialog for input)
*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.2.5";

// =========================================
// ユーザー設定 / User settings
// =========================================
/* 「選択後に移動」先のレイヤー名 / Destination layer name for "move after selection" */
var TEXT_LAYER_NAME = "_text";

// =========================================
// ローカライズ / Localization
// =========================================
var currentLanguage = ($.locale || "").toLowerCase().indexOf("ja") === 0 ? "ja" : "en";

var LABELS = {
    dialog: {
        title: { ja: "テキストを選択", en: "Select Text" },
        bulkEditTitle: { ja: "テキストを一括変更", en: "Bulk Edit Text" }
    },

    error: {
        noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
        emptyKeyword: { ja: "検索文字列を入力してください。", en: "Enter a search string." },
        invalidRegex: { ja: "正規表現が正しくありません。", en: "The regular expression is invalid." },
        noMatch: { ja: "条件に一致するテキストが見つかりませんでした。", en: "No text matched the criteria." }
    },

    panel: {
        selection: { ja: "選択条件", en: "Selection Criteria" },
        artboardScope: { ja: "対象とするアートボード", en: "Target Artboard" },
        attribute: { ja: "属性で選択", en: "Select by Attribute" },
        textType: { ja: "テキストの種類", en: "Text Type" },
        textMatch: { ja: "文字列で選択", en: "Select by String" },
        postProcess: { ja: "選択後の処理", en: "After Selection" }
    },

    attribute: {
        fontFamily: { ja: "フォントファミリー", en: "Font Family" },
        fontFamilyStyle: { ja: "+ スタイル", en: "+ Style" },
        fontFamilyStyleSize: { ja: "+ スタイルとサイズ", en: "+ Style and Size" },
        fontSize: { ja: "フォントサイズ", en: "Font Size" },
        textFillColor: { ja: "テキストカラー", en: "Text Color" },
        opacity: { ja: "不透明度", en: "Opacity" }
    },

    artboard: {
        all: { ja: "すべて", en: "All" },
        current: { ja: "現在のアートボードのみ", en: "Current Artboard Only" }
    },

    textType: {
        all: { ja: "すべて", en: "All" },
        point: { ja: "ポイント文字", en: "Point Text" },
        area: { ja: "エリア内文字", en: "Area Text" },
        path: { ja: "パス上文字", en: "Path Text" }
    },

    match: {
        exact: { ja: "完全一致", en: "Exact" },
        contains: { ja: "部分一致", en: "Contains" },
        startsWith: { ja: "先頭一致", en: "Starts With" },
        endsWith: { ja: "末尾一致", en: "Ends With" },
        regex: { ja: "正規表現", en: "Regex" }
    },

    postProcess: {
        none: { ja: "なし", en: "None" },
        hide: { ja: "選択後に非表示", en: "Hide after selection" },
        hideOthers: { ja: "選択したテキスト以外を非表示", en: "Hide all but selected" },
        moveToTextLayer: { ja: "選択後に「_text」レイヤーへ移動", en: "Move to “_text” layer after selection" },
        bulkEdit: { ja: "一括編集", en: "Bulk edit" }
    },

    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },

    tip: {
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
        pointText: {
            ja: "ポイント文字だけを選択します。ショートカット：Option + W",
            en: "Select point text only. Shortcut: Option + W"
        },
        areaText: {
            ja: "エリア内文字だけを選択します。ショートカット：Option + E",
            en: "Select area text only. Shortcut: Option + E"
        },
        pathText: {
            ja: "パス上文字だけを選択します。",
            en: "Select path text only."
        },
        keywordInput: {
            ja: "検索対象にする文字列を入力します。空欄では実行できません。",
            en: "Enter the search string. This cannot be empty."
        },
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
        artboardAll: {
            ja: "ドキュメント内のすべてのアートボードを対象に選択します。",
            en: "Select across all artboards in the document."
        },
        artboardCurrent: {
            ja: "現在のアートボードに重なるオブジェクトだけを対象に選択します。",
            en: "Select only objects that overlap the active artboard."
        },
        noPostProcess: {
            ja: "選択後に追加処理を行いません。",
            en: "Do not apply any additional processing after selection."
        },
        hideAfterSelection: {
            ja: "選択されたテキストを非表示にします。",
            en: "Hide the selected text."
        },
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
    }
};

function L(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node && node[parts[i]] != null) {
            node = node[parts[i]];
        } else {
            return key;
        }
    }
    if (node && node[currentLanguage]) {
        return node[currentLanguage];
    }
    if (node && node.en) {
        return node.en;
    }
    return key;
}

(function () {
    if (app.documents.length === 0) {
        alert(L("error.noDocument"));
        return;
    }

    // =========================================
    // UI共通設定 / Shared UI settings
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];

    /* パネル共通レイアウトを設定 / Setup shared panel layout */
    function setupPanelLayout(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    /* ヘルプチップを設定 / Set help tip text */
    function setHelpTip(control, text) {
        if (control && text) {
            control.helpTip = text;
        }
    }

    /* 選択内容を配列化 / Convert selection value to array */
    function collectSelectionAsArray(selection) {
        var selectedItems = [];
        if (!selection) {
            return selectedItems;
        }

        if (selection.typename) {
            selectedItems.push(selection);
            return selectedItems;
        }

        if (typeof selection.length === "number") {
            for (var i = 0; i < selection.length; i++) {
                selectedItems.push(selection[i]);
            }
            return selectedItems;
        }

        selectedItems.push(selection);
        return selectedItems;
    }

    /* 選択内容からテキスト範囲を取得 / Get text range from selected item */
    function getTextRangeFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return null;
        }

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

    /* 選択内容から文字列を取得 / Get text contents from selected item */
    function getTextContentsFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return "";
        }

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

    // =========================================
    // テキスト取得 / Text retrieval
    // =========================================
    /* 選択中テキストの文字列を取得 / Get text contents from selected text frame or range */
    function getSelectedTextString() {
        var selectedItems = collectSelectionAsArray(app.activeDocument.selection);
        if (selectedItems.length === 0) {
            return "";
        }

        for (var i = 0; i < selectedItems.length; i++) {
            var textContents = getTextContentsFromSelectionItem(selectedItems[i]);
            if (textContents) {
                return textContents;
            }
        }
        return "";
    }

    /* 属性プレビューの初期値（未取得を表す「—」） / Empty attribute preview ("—" means not available) */
    function makeEmptyAttributePreview() {
        return {
            family: "—",
            familyStyle: "—",
            familyStyleSize: "—",
            size: "—",
            opacity: "—"
        };
    }

    /* 選択中テキストの文字属性を取得 / Get text attributes from selected text frame or range */
    function getSelectedTextAttributePreview() {
        var selectedItems = collectSelectionAsArray(app.activeDocument.selection);
        if (selectedItems.length === 0) {
            return makeEmptyAttributePreview();
        }

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

    /* テキスト範囲から文字属性プレビューを生成 / Build text attribute preview from text range */
    function getTextAttributePreviewFromTextRange(textRange) {
        var preview = makeEmptyAttributePreview();

        try {
            var characterAttributes = textRange.characterAttributes;
            var textFont = characterAttributes.textFont;
            var fontSize = characterAttributes.size;
            var familyName = textFont.family || "";
            var styleName = textFont.style || "";
            var sizeText = fontSize ? String(Math.round(fontSize * 100) / 100) + " pt" : "—";

            preview.family = familyName || "—";
            preview.familyStyle = familyName ? familyName + (styleName ? " / " + styleName : "") : "—";
            preview.familyStyleSize = preview.familyStyle !== "—" ? preview.familyStyle + " / " + sizeText : "—";
            preview.size = sizeText;
        } catch (attributePreviewError) {
            return preview;
        }

        return preview;
    }

    /* 選択内容から不透明度プレビューを生成 / Build opacity preview from selected item */
    function getOpacityPreviewFromSelectionItem(selectionItem) {
        try {
            if (selectionItem && typeof selectionItem.opacity === "number") {
                return String(Math.round(selectionItem.opacity * 100) / 100) + " %";
            }
        } catch (opacityPreviewError) {
        }
        return "—";
    }

    /* ラジオボタンとプレビュー値の行を作成 / Create radio button row with preview value */
    function addRadioWithPreview(parent, label, previewText, labelWidth) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = "fill";

        var radioButton = rowGroup.add("radiobutton", undefined, label);
        if (typeof labelWidth === "number") {
            radioButton.preferredSize.width = labelWidth;
        }
        var previewLabel = rowGroup.add("statictext", undefined, previewText);
        previewLabel.alignment = ["fill", "center"];
        previewLabel.characters = 24;

        return {
            radioButton: radioButton,
            previewLabel: previewLabel
        };
    }

    /* ラジオボタンのみの行を作成 / Create radio button row without preview value */
    function addRadioOnlyRow(parent, label, labelWidth) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = "fill";

        var radioButton = rowGroup.add("radiobutton", undefined, label);
        if (typeof labelWidth === "number") {
            radioButton.preferredSize.width = labelWidth;
        }

        return {
            radioButton: radioButton,
            previewLabel: null
        };
    }

    var initialString = getSelectedTextString();
    var hasInitialSelection = collectSelectionAsArray(app.activeDocument.selection).length > 0;
    var initialAttributePreview = getSelectedTextAttributePreview();

    // =========================================
    // ダイアログボックス / Dialog box
    // =========================================

    var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);

    /* 選択条件パネルを作成 / Create selection criteria panel */
    var selectionPanel = dialog.add("panel", undefined, L("panel.selection"));
    setupPanelLayout(selectionPanel, 10);
    selectionPanel.margins = [15, 20, 15, 15];

    /* 対象アートボードパネル / Target artboard panel */
    var artboardScopePanel = selectionPanel.add("panel", undefined, L("panel.artboardScope"));
    artboardScopePanel.orientation = "row";
    artboardScopePanel.alignChildren = ["left", "center"];
    artboardScopePanel.alignment = "fill";
    artboardScopePanel.margins = [15, 20, 15, 15];
    artboardScopePanel.spacing = 20;

    var rbArtboardAll = artboardScopePanel.add("radiobutton", undefined, L("artboard.all"));
    var rbArtboardCurrent = artboardScopePanel.add("radiobutton", undefined, L("artboard.current"));
    rbArtboardAll.value = true;

    setHelpTip(rbArtboardAll, L("tip.artboardAll"));
    setHelpTip(rbArtboardCurrent, L("tip.artboardCurrent"));

    setupExclusiveRadioButtons([rbArtboardAll, rbArtboardCurrent]);

    /* 属性選択パネル / Attribute selection panel */
    var attributePanel = selectionPanel.add("panel", undefined, L("panel.attribute"));
    setupPanelLayout(attributePanel, 6);
    attributePanel.enabled = hasInitialSelection;

    var ATTRIBUTE_LABEL_WIDTH = 150;
    var fontFamilyPreviewRow = addRadioWithPreview(attributePanel, L("attribute.fontFamily"), initialAttributePreview.family, ATTRIBUTE_LABEL_WIDTH);
    var fontFamilyStylePreviewRow = addRadioWithPreview(attributePanel, L("attribute.fontFamilyStyle"), initialAttributePreview.familyStyle, ATTRIBUTE_LABEL_WIDTH);
    var fontFamilyStyleSizePreviewRow = addRadioWithPreview(attributePanel, L("attribute.fontFamilyStyleSize"), initialAttributePreview.familyStyleSize, ATTRIBUTE_LABEL_WIDTH);
    var fontSizePreviewRow = addRadioWithPreview(attributePanel, L("attribute.fontSize"), initialAttributePreview.size, ATTRIBUTE_LABEL_WIDTH);

    var rbFontFamily = fontFamilyPreviewRow.radioButton;
    var rbFontFamilyStyle = fontFamilyStylePreviewRow.radioButton;
    var rbFontFamilyStyleSize = fontFamilyStyleSizePreviewRow.radioButton;
    var rbFontSize = fontSizePreviewRow.radioButton;
    var rbTextFillColor = addRadioOnlyRow(attributePanel, L("attribute.textFillColor"), ATTRIBUTE_LABEL_WIDTH).radioButton;
    var opacityPreviewRow = addRadioWithPreview(attributePanel, L("attribute.opacity"), initialAttributePreview.opacity, ATTRIBUTE_LABEL_WIDTH);
    var rbOpacity = opacityPreviewRow.radioButton;

    setHelpTip(rbFontFamily, L("tip.fontFamily"));
    setHelpTip(rbFontFamilyStyle, L("tip.fontFamilyStyle"));
    setHelpTip(rbFontFamilyStyleSize, L("tip.fontFamilyStyleSize"));
    setHelpTip(rbFontSize, L("tip.fontSize"));
    setHelpTip(rbTextFillColor, L("tip.textFillColor"));
    setHelpTip(rbOpacity, L("tip.opacity"));

    /* テキスト種類と文字列条件を横並びに配置 / Arrange text type and string condition panels side by side */
    var textConditionGroup = selectionPanel.add("group");
    textConditionGroup.orientation = "row";
    textConditionGroup.alignChildren = "fill";
    textConditionGroup.alignment = "fill";

    /* テキスト種類パネルを追加 / Add text type panel */
    var textTypePanel = textConditionGroup.add("panel", undefined, L("panel.textType"));
    setupPanelLayout(textTypePanel, 6);

    /* テキスト種類ラジオボタン / Text type radio buttons */
    var rbAllText = textTypePanel.add("radiobutton", undefined, L("textType.all"));
    var rbPointText = textTypePanel.add("radiobutton", undefined, L("textType.point"));
    var rbAreaText = textTypePanel.add("radiobutton", undefined, L("textType.area"));
    var rbPathText = textTypePanel.add("radiobutton", undefined, L("textType.path"));

    setHelpTip(rbAllText, L("tip.allText"));
    setHelpTip(rbPointText, L("tip.pointText"));
    setHelpTip(rbAreaText, L("tip.areaText"));
    setHelpTip(rbPathText, L("tip.pathText"));

    /* 初期選択を設定（選択があれば「＋スタイルとサイズ」、なければ「すべて」） / Set initial selection */
    if (hasInitialSelection) {
        rbFontFamilyStyleSize.value = true;
    } else {
        rbAllText.value = true;
    }

    /* 文字列条件パネルを作成 / Create string condition panel */
    var textMatchPanelTitle = L("panel.textMatch");
    var textMatchPanel = textConditionGroup.add("panel", undefined, textMatchPanelTitle);
    setupPanelLayout(textMatchPanel, 6);

    var textKeywordInput = textMatchPanel.add("edittext", undefined, initialString);
    textKeywordInput.characters = 22;
    setHelpTip(textKeywordInput, L("tip.keywordInput"));

    var textMatchOptionsGroup = textMatchPanel.add("group");
    textMatchOptionsGroup.orientation = "row";
    textMatchOptionsGroup.alignChildren = "top";
    textMatchOptionsGroup.margins = [0, 10, 0, 0];
    textMatchOptionsGroup.spacing = 22;

    var textMatchLeftColumnGroup = textMatchOptionsGroup.add("group");
    textMatchLeftColumnGroup.orientation = "column";
    textMatchLeftColumnGroup.alignChildren = "left";

    var textMatchCenterColumnGroup = textMatchOptionsGroup.add("group");
    textMatchCenterColumnGroup.orientation = "column";
    textMatchCenterColumnGroup.alignChildren = "left";

    var textMatchRightColumnGroup = textMatchOptionsGroup.add("group");
    textMatchRightColumnGroup.orientation = "column";
    textMatchRightColumnGroup.alignChildren = "left";

    var rbExactMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, L("match.exact"));
    var rbLooseMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, L("match.contains"));

    var rbStartsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, L("match.startsWith"));
    var rbEndsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, L("match.endsWith"));

    var rbRegexMatch = textMatchRightColumnGroup.add("radiobutton", undefined, L("match.regex"));

    setHelpTip(rbExactMatch, L("tip.exactMatch"));
    setHelpTip(rbStartsWith, L("tip.startsWith"));
    setHelpTip(rbEndsWith, L("tip.endsWith"));
    setHelpTip(rbLooseMatch, L("tip.containsMatch"));
    setHelpTip(rbRegexMatch, L("tip.regexMatch"));

    /* 選択条件ラジオボタンをまとめる / Collect selection condition radio buttons */
    var selectionRadios = [
        rbFontFamily, rbFontFamilyStyle, rbFontFamilyStyleSize, rbFontSize, rbTextFillColor, rbOpacity,
        rbAllText, rbPointText, rbAreaText, rbPathText,
        rbExactMatch, rbStartsWith, rbEndsWith, rbLooseMatch, rbRegexMatch
    ];

    // =========================================
    // ラジオボタン制御 / Radio button control
    // =========================================
    /* ラジオボタンを排他化 / Make radio buttons mutually exclusive */
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

    setupExclusiveRadioButtons(selectionRadios);

    /* 指定したラジオボタンだけを選択 / Select only the specified radio button */
    function selectExclusiveRadioButton(radioButtons, targetRadioButton) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (radioButtons[i] === targetRadioButton);
        }
    }

    /* テキスト入力欄にフォーカスがあるか判定 / Check whether an edit text field has focus */
    function isEditTextFocused(editText) {
        try {
            return editText && editText.active;
        } catch (focusCheckError) {
            return false;
        }
    }

    /* キー入力で選択条件を切り替え / Switch selection condition by keyboard */
    function addSelectionKeyHandler(dialog, radioButtons) {
        dialog.addEventListener("keydown", function (event) {
            if (isEditTextFocused(textKeywordInput)) {
                return;
            }
            if (!event.altKey) {
                return;
            }

            var targetRadioButton = null;

            if (event.keyName === "Q") {
                targetRadioButton = rbAllText;
            } else if (event.keyName === "W") {
                targetRadioButton = rbPointText;
            } else if (event.keyName === "E") {
                targetRadioButton = rbAreaText;
            } else if (event.keyName === "A") {
                targetRadioButton = rbExactMatch;
            } else if (event.keyName === "B") {
                targetRadioButton = rbStartsWith;
            } else if (event.keyName === "D") {
                targetRadioButton = rbEndsWith;
            } else if (event.keyName === "I") {
                targetRadioButton = rbLooseMatch;
            } else if (event.keyName === "R") {
                targetRadioButton = rbRegexMatch;
            }

            if (targetRadioButton) {
                selectExclusiveRadioButton(radioButtons, targetRadioButton);
                event.preventDefault();
            }
        });
    }

    addSelectionKeyHandler(dialog, selectionRadios);

    /* 属性ラジオボタン用コマンド関数 / Command helpers for attribute radio buttons */
    /* 選択中属性検索コマンドを取得 / Get selected attribute search command */
    function getSelectedAttributeCommand() {
        if (rbFontFamily.value) {
            return "Find Text Font Family menu item";
        }
        if (rbFontFamilyStyle.value) {
            return "Find Text Font Family Style menu item";
        }
        if (rbFontFamilyStyleSize.value) {
            return "Find Text Font Family Style Size menu item";
        }
        if (rbFontSize.value) {
            return "Find Text Font Size menu item";
        }
        if (rbTextFillColor.value) {
            return "Find Text Fill Color menu item";
        }
        if (rbOpacity.value) {
            return "Find Opacity menu item";
        }
        return "";
    }

    /* 属性検索コマンドを実行 / Execute attribute search command */
    function executeSelectedAttributeCommand(commandName) {
        if (!commandName) {
            return false;
        }
        app.executeMenuCommand(commandName);
        return true;
    }

    /* 文字列一致モードを取得 / Get selected text match mode */
    function getSelectedTextMatchMode() {
        if (rbExactMatch.value) {
            return "exact";
        }
        if (rbStartsWith.value) {
            return "startsWith";
        }
        if (rbEndsWith.value) {
            return "endsWith";
        }
        if (rbLooseMatch.value) {
            return "contains";
        }
        if (rbRegexMatch.value) {
            return "regex";
        }
        return "";
    }

    /* 文字列一致を判定 / Check whether text matches keyword */
    function textMatches(text, keyword, matchMode, regex) {
        if (matchMode === "regex") {
            return regex ? regex.test(text) : false;
        }
        if (matchMode === "exact") {
            return text === keyword;
        }
        if (matchMode === "startsWith") {
            return text.indexOf(keyword) === 0;
        }
        if (matchMode === "endsWith") {
            return text.lastIndexOf(keyword) === text.length - keyword.length;
        }
        if (matchMode === "contains") {
            return text.indexOf(keyword) !== -1;
        }
        return false;
    }

    /* 文字列条件の入力を検証 / Validate string condition input */
    function validateTextMatchInput(keyword, matchMode) {
        if (!keyword) {
            alert(L("error.emptyKeyword"));
            return null;
        }
        if (matchMode === "regex") {
            try {
                return { regex: new RegExp(keyword) };
            } catch (regexError) {
                alert(L("error.invalidRegex"));
                return null;
            }
        }
        return { regex: null };
    }

    /* 述語に一致するテキストフレームを選択 / Select text frames matching a predicate */
    function selectTextFrames(predicate, artboardScope) {
        var doc = app.activeDocument;
        var allItems = doc.textFrames;
        var selected = [];

        for (var i = 0; i < allItems.length; i++) {
            if (predicate(allItems[i])) {
                selected.push(allItems[i]);
            }
        }

        if (artboardScope === "current") {
            selected = filterItemsByActiveArtboard(selected);
        }

        doc.selection = selected;
        return selected.length;
    }

    /* 選択中のテキスト種類を取得 / Get selected text type */
    function getSelectedTextType() {
        if (rbAllText.value) {
            return "all";
        }
        if (rbPointText.value) {
            return "point";
        }
        if (rbAreaText.value) {
            return "area";
        }
        if (rbPathText.value) {
            return "path";
        }
        return "all";
    }

    /* 選択後の処理パネルを作成 / Create after-selection processing panel */
    var optionPanel = dialog.add("panel", undefined, L("panel.postProcess"));
    setupPanelLayout(optionPanel, 6);
    /* ドキュメント内に TextFrame が 0 件なら後処理は無意味なのでディム / Disable post-process when document has no text frames */
    optionPanel.enabled = app.activeDocument.textFrames.length > 0;

    var rbNoPostProcess = optionPanel.add("radiobutton", undefined, L("postProcess.none"));
    var rbHide = optionPanel.add("radiobutton", undefined, L("postProcess.hide"));
    var rbHideOthers = optionPanel.add("radiobutton", undefined, L("postProcess.hideOthers"));
    var rbMove = optionPanel.add("radiobutton", undefined, L("postProcess.moveToTextLayer"));
    var rbBulkEdit = optionPanel.add("radiobutton", undefined, L("postProcess.bulkEdit"));

    rbNoPostProcess.value = true;

    setHelpTip(rbNoPostProcess, L("tip.noPostProcess"));
    setHelpTip(rbHide, L("tip.hideAfterSelection"));
    setHelpTip(rbHideOthers, L("tip.hideOthers"));
    setHelpTip(rbMove, L("tip.moveToTextLayer"));
    setHelpTip(rbBulkEdit, L("tip.bulkEdit"));

    setupExclusiveRadioButtons([
        rbNoPostProcess,
        rbHide,
        rbHideOthers,
        rbMove,
        rbBulkEdit
    ]);

    // =========================================
    // 後処理 / Post-processing
    // =========================================
    /* 選択中の後処理モードを取得 / Get selected post-process mode */
    function getSelectedPostProcessMode() {
        if (rbNoPostProcess.value) {
            return "";
        }
        if (rbHide.value) {
            return "hide";
        }
        if (rbHideOthers.value) {
            return "hideOthers";
        }
        if (rbMove.value) {
            return "moveToTextLayer";
        }
        if (rbBulkEdit.value) {
            return "bulkEdit";
        }
        return "";
    }

    /* テキストフレームを一括編集 / Bulk edit selected text frames */
    function bulkEditTextFrames(textFrames) {
        if (textFrames.length === 0) {
            return;
        }

        var bulkDialog = new Window("dialog", L("dialog.bulkEditTitle"));
        bulkDialog.orientation = "column";
        bulkDialog.alignChildren = "fill";
        bulkDialog.spacing = 10;
        bulkDialog.margins = 16;

        var inputText = bulkDialog.add("edittext", undefined, "", {
            multiline: true,
            scrolling: true
        });
        inputText.preferredSize = [300, 70];
        inputText.active = true;

        var bulkButtonGroup = bulkDialog.add("group");
        bulkButtonGroup.alignment = "right";
        var bulkCancelBtn = bulkButtonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var bulkOkBtn = bulkButtonGroup.add("button", undefined, "OK", { name: "ok" });

        bulkOkBtn.onClick = function () {
            bulkDialog.close(1);
        };
        bulkCancelBtn.onClick = function () {
            bulkDialog.close(0);
        };

        if (bulkDialog.show() !== 1) {
            return;
        }

        var newText = inputText.text;
        for (var i = 0; i < textFrames.length; i++) {
            var tf = textFrames[i];
            var originalRange = tf.textRange;
            var originalLength = originalRange.length;
            if (originalLength === 0) {
                continue;
            }

            tf.contents = newText;

            var newRange = tf.textRange;
            var minLen = Math.min(originalLength, newRange.length);
            for (var j = 0; j < minLen; j++) {
                try {
                    newRange.characters[j].characterAttributes = originalRange.characters[j].characterAttributes;
                } catch (attrCopyError) {
                }
            }
        }

        app.redraw();
    }

    /* レイヤーを取得または作成（元の状態を退避） / Get or create layer, capturing original state */
    function getOrCreateLayerByName(doc, layerName) {
        var targetLayer;
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

    /* 選択中の子孫を持つか判定 / Check whether the item contains any selected descendant */
    function containsSelectedDescendant(item) {
        var children;
        try {
            children = item.pageItems;
        } catch (childAccessError) {
            return false;
        }
        if (!children || children.length === 0) {
            return false;
        }
        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            try {
                if (child.selected) {
                    return true;
                }
            } catch (selectedReadError) {
            }
            if (containsSelectedDescendant(child)) {
                return true;
            }
        }
        return false;
    }

    /* 選択中以外のオブジェクトを非表示（ロックは除く） / Hide all objects except the selection (locked items are skipped) */
    function hideItemsExceptSelection(doc) {
        var allPageItems = doc.pageItems;
        for (var i = 0; i < allPageItems.length; i++) {
            var pageItem = allPageItems[i];
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
                /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
            }
        }
    }

    /* 選択後の後処理を適用 / Apply post-process to selected items */
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
            for (var i = 0; i < selectedItems.length; i++) {
                try {
                    selectedItems[i].hidden = true;
                } catch (hideItemError) {
                    /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
                }
            }
            return;
        }

        if (postProcessMode === "hideOthers") {
            hideItemsExceptSelection(doc);
            return;
        }

        if (postProcessMode === "bulkEdit") {
            var textFramesOnly = [];
            for (var k = 0; k < selectedItems.length; k++) {
                if (selectedItems[k].typename === "TextFrame") {
                    textFramesOnly.push(selectedItems[k]);
                }
            }
            bulkEditTextFrames(textFramesOnly);
            return;
        }

        if (postProcessMode === "moveToTextLayer") {
            var layerInfo = getOrCreateLayerByName(doc, TEXT_LAYER_NAME);
            var textLayer = layerInfo.layer;
            for (var j = 0; j < selectedItems.length; j++) {
                try {
                    selectedItems[j].locked = false;
                    selectedItems[j].move(textLayer, ElementPlacement.PLACEATBEGINNING);
                } catch (moveItemError) {
                    /* ロック状態や親レイヤーの状態によって移動できない場合はスキップ / Skip items that cannot be moved because of lock state or parent layer state */
                }
            }
            /* 元のロック・可視状態へ戻す（新規作成時はデフォルト値の戻りで実質 no-op） / Restore original lock/visibility (no-op when newly created) */
            try {
                textLayer.locked = layerInfo.originalLocked;
                textLayer.visible = layerInfo.originalVisible;
            } catch (restoreLayerError) {
            }
        }
    }

    // =========================================
    // 対象アートボードによる絞り込み / Filter by target artboard
    // =========================================
    /* 対象アートボードスコープを取得 / Get target artboard scope */
    function getSelectedArtboardScope() {
        return rbArtboardCurrent.value ? "current" : "all";
    }

    /* 現在のアートボードの矩形を取得 / Get active artboard rectangle */
    function getActiveArtboardRect(doc) {
        try {
            var activeIndex = doc.artboards.getActiveArtboardIndex();
            return doc.artboards[activeIndex].artboardRect; /* [left, top, right, bottom] */
        } catch (artboardRectError) {
            return null;
        }
    }

    /* アイテムが矩形に重なるか判定 / Check whether item overlaps the rectangle */
    function isItemWithinArtboardRect(item, artboardRect) {
        if (!artboardRect) {
            return true;
        }
        try {
            var bounds = item.visibleBounds; /* [left, top, right, bottom] */
            var overlapsHorizontally = bounds[2] >= artboardRect[0] && bounds[0] <= artboardRect[2];
            var overlapsVertically = bounds[1] >= artboardRect[3] && bounds[3] <= artboardRect[1];
            return overlapsHorizontally && overlapsVertically;
        } catch (boundsError) {
            return true;
        }
    }

    /* 現在のアートボード内のアイテムへ絞り込み / Filter items to those on the active artboard */
    function filterItemsByActiveArtboard(items) {
        var artboardRect = getActiveArtboardRect(app.activeDocument);
        var filtered = [];
        for (var i = 0; i < items.length; i++) {
            if (isItemWithinArtboardRect(items[i], artboardRect)) {
                filtered.push(items[i]);
            }
        }
        return filtered;
    }

    // =========================================
    // テキスト選択処理 / Text selection processing
    // =========================================
    /* テキスト種類に対応する述語を生成 / Build a predicate for a text type */
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

    // =========================================
    // ボタンエリア / Button area
    // =========================================
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelBtn = buttonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    /* 選択件数を確定し後処理を適用（0件ならアラート） / Finalize by count, then apply post-process (alert if none) */
    function finalizeSelection(count, postProcessMode) {
        if (count === 0) {
            alert(L("error.noMatch"));
            return;
        }
        applyPostProcessToSelection(postProcessMode);
    }

    /* OKボタン実行処理 / Handle OK button action */
    okBtn.onClick = function () {
        var selectedPostProcessMode = getSelectedPostProcessMode();
        var artboardScope = getSelectedArtboardScope();

        /* 属性で選択：Illustrator標準コマンドに選択を任せる / By attribute: let Illustrator's command perform the selection */
        var selectedAttributeCommand = getSelectedAttributeCommand();
        if (selectedAttributeCommand) {
            dialog.close();
            executeSelectedAttributeCommand(selectedAttributeCommand);
            if (artboardScope === "current") {
                var filteredAttributeSelection = filterItemsByActiveArtboard(collectSelectionAsArray(app.activeDocument.selection));
                app.activeDocument.selection = filteredAttributeSelection;
                finalizeSelection(filteredAttributeSelection.length, selectedPostProcessMode);
                return;
            }
            applyPostProcessToSelection(selectedPostProcessMode);
            return;
        }

        /* 文字列で選択 / By string */
        var textMatchMode = getSelectedTextMatchMode();
        if (textMatchMode) {
            var keyword = textKeywordInput.text || "";
            var validation = validateTextMatchInput(keyword, textMatchMode);
            if (!validation) {
                return;
            }
            dialog.close();
            var stringPredicate = function (textFrame) {
                return textMatches(textFrame.contents || "", keyword, textMatchMode, validation.regex);
            };
            finalizeSelection(selectTextFrames(stringPredicate, artboardScope), selectedPostProcessMode);
            return;
        }

        /* テキストの種類で選択 / By text type */
        dialog.close();
        var typePredicate = buildTextTypePredicate(getSelectedTextType());
        finalizeSelection(selectTextFrames(typePredicate, artboardScope), selectedPostProcessMode);
    };

    /* キャンセルボタン処理 / Handle Cancel button action */
    cancelBtn.onClick = function () {
        dialog.close();
    };

    /* ダイアログを表示 / Show dialog */
    dialog.show();
}());