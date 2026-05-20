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
// バージョンとローカライズ
// =========================================
var SCRIPT_VERSION = "v1.0.1";

var currentLanguage = ($.locale || "").toLowerCase().indexOf("ja") === 0 ? "ja" : "en";

var LABELS = {
    dialogTitle: { ja: "テキストを選択", en: "Select Text" },
    errNoDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
    errEmptyKeyword: { ja: "検索文字列を入力してください。", en: "Enter a search string." },
    errInvalidRegex: { ja: "正規表現が正しくありません。", en: "The regular expression is invalid." },
    errNoMatch: { ja: "条件に一致するテキストが見つかりませんでした。", en: "No text matched the criteria." },

    selectionPanel: { ja: "選択条件", en: "Selection Criteria" },
    attributePanel: { ja: "属性で選択", en: "Select by Attribute" },
    textTypePanel: { ja: "テキストの種類", en: "Text Type" },
    textMatchPanel: { ja: "文字列で選択", en: "Select by String" },
    postProcessPanel: { ja: "選択後の処理", en: "After Selection" },

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
    moveToTextLayer: { ja: "選択後に「_text」レイヤーへ移動", en: "Move to “_text” layer after selection" },
    bulkEdit: { ja: "一括編集", en: "Bulk edit" },
    bulkEditDialogTitle: { ja: "テキストを一括変更", en: "Bulk Edit Text" },
    cancel: { ja: "キャンセル", en: "Cancel" },

    tipFontFamily: { ja: "選択中のテキストと同じフォントファミリーのテキストを、Illustrator標準コマンドで検索します。", en: "Find text with the same font family as the selected text using Illustrator’s built-in command." },
    tipFontFamilyStyle: { ja: "選択中のテキストと同じフォントファミリー＋スタイルのテキストを、Illustrator標準コマンドで検索します。", en: "Find text with the same font family and style as the selected text using Illustrator’s built-in command." },
    tipFontFamilyStyleSize: { ja: "選択中のテキストと同じフォントファミリー＋スタイル＋サイズのテキストを、Illustrator標準コマンドで検索します。", en: "Find text with the same font family, style, and size as the selected text using Illustrator’s built-in command." },
    tipFontSize: { ja: "選択中のテキストと同じフォントサイズのテキストを、Illustrator標準コマンドで検索します。", en: "Find text with the same font size as the selected text using Illustrator’s built-in command." },
    tipTextFillColor: { ja: "選択中のテキストと同じテキストカラーのテキストを、Illustrator標準コマンドで検索します。", en: "Find text with the same text color as the selected text using Illustrator’s built-in command." },
    tipOpacity: { ja: "選択中のオブジェクトと同じ不透明度のオブジェクトを、Illustrator標準コマンドで検索します。テキスト以外のオブジェクトも対象になります。", en: "Find objects with the same opacity as the selected object using Illustrator’s built-in command. Non-text objects are also matched." },

    tipAllText: { ja: "ドキュメント内のすべてのテキストフレームを選択します。ショートカット：Option + Q", en: "Select all text frames in the document. Shortcut: Option + Q" },
    tipPointText: { ja: "ポイント文字だけを選択します。ショートカット：Option + W", en: "Select point text only. Shortcut: Option + W" },
    tipAreaText: { ja: "エリア内文字だけを選択します。ショートカット：Option + E", en: "Select area text only. Shortcut: Option + E" },
    tipPathText: { ja: "パス上文字だけを選択します。", en: "Select path text only." },
    tipKeywordInput: { ja: "検索対象にする文字列を入力します。空欄では実行できません。", en: "Enter the search string. This cannot be empty." },
    tipExactMatch: { ja: "入力した文字列と完全に一致するテキストを選択します。大小区別あり。ショートカット：Option + A", en: "Select text that exactly matches the entered string. Case-sensitive. Shortcut: Option + A" },
    tipStartsWith: { ja: "入力した文字列で始まるテキストを選択します。大小区別あり。ショートカット：Option + B", en: "Select text that starts with the entered string. Case-sensitive. Shortcut: Option + B" },
    tipEndsWith: { ja: "入力した文字列で終わるテキストを選択します。大小区別あり。ショートカット：Option + D", en: "Select text that ends with the entered string. Case-sensitive. Shortcut: Option + D" },
    tipContainsMatch: { ja: "入力した文字列を含むテキストを選択します。大小区別あり。ショートカット：Option + I", en: "Select text that contains the entered string. Case-sensitive. Shortcut: Option + I" },
    tipRegexMatch: { ja: "入力した正規表現に一致するテキストを選択します。大小区別あり。ショートカット：Option + R", en: "Select text that matches the entered regular expression. Case-sensitive. Shortcut: Option + R" },

    tipNoPostProcess: { ja: "選択後に追加処理を行いません。", en: "Do not apply any additional processing after selection." },
    tipHideAfterSelection: { ja: "選択されたテキストを非表示にします。", en: "Hide the selected text." },
    tipMoveToTextLayer: { ja: "選択されたテキストを「_text」レイヤーへ移動します。レイヤーがない場合は作成します。", en: "Move the selected text to the “_text” layer. The layer will be created if it does not exist." },
    tipBulkEdit: { ja: "選択されたテキストの内容をまとめて置き換えます。書式は維持されます。", en: "Replace contents of all selected text frames at once. Formatting is preserved." }
};

function L(key) {
    if (LABELS[key] && LABELS[key][currentLanguage]) {
        return LABELS[key][currentLanguage];
    }
    if (LABELS[key] && LABELS[key].en) {
        return LABELS[key].en;
    }
    return key;
}

(function () {
    if (app.documents.length === 0) {
        alert(L("errNoDocument"));
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

        try {
            if (selection.typename) {
                selectedItems.push(selection);
                return selectedItems;
            }
        } catch (selectionTypeError) {
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

        try {
            var textRange = getTextRangeFromSelectionItem(selectionItem);
            if (textRange && typeof textRange.contents === "string") {
                return textRange.contents;
            }
        } catch (textRangeContentsError) {
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

    /* 選択中テキストの文字属性を取得 / Get text attributes from selected text frame or range */
    function getSelectedTextAttributePreview() {
        var preview = {
            family: "—",
            familyStyle: "—",
            familyStyleSize: "—",
            size: "—",
            opacity: "—"
        };

        var selectedItems = collectSelectionAsArray(app.activeDocument.selection);
        if (selectedItems.length === 0) {
            return preview;
        }

        for (var i = 0; i < selectedItems.length; i++) {
            var textRange = getTextRangeFromSelectionItem(selectedItems[i]);
            if (textRange) {
                var rangePreview = getTextAttributePreviewFromTextRange(textRange);
                rangePreview.opacity = getOpacityPreviewFromSelectionItem(selectedItems[i]);
                return rangePreview;
            }
        }

        return preview;
    }

    /* テキスト範囲から文字属性プレビューを生成 / Build text attribute preview from text range */
    function getTextAttributePreviewFromTextRange(textRange) {
        var preview = {
            family: "—",
            familyStyle: "—",
            familyStyleSize: "—",
            size: "—",
            opacity: "—"
        };

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

    var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);

    /* 選択条件パネルを作成 / Create selection criteria panel */
    var selectionPanel = dialog.add("panel", undefined, L("selectionPanel"));
    setupPanelLayout(selectionPanel, 10);
    selectionPanel.margins = [15, 20, 15, 15];

    /* 属性選択パネル / Attribute selection panel */
    var attributePanel = selectionPanel.add("panel", undefined, L("attributePanel"));
    setupPanelLayout(attributePanel, 6);
    attributePanel.enabled = hasInitialSelection;

    var ATTRIBUTE_LABEL_WIDTH = 150;
    var fontFamilyPreviewRow = addRadioWithPreview(attributePanel, L("fontFamily"), initialAttributePreview.family, ATTRIBUTE_LABEL_WIDTH);
    var fontFamilyStylePreviewRow = addRadioWithPreview(attributePanel, L("fontFamilyStyle"), initialAttributePreview.familyStyle, ATTRIBUTE_LABEL_WIDTH);
    var fontFamilyStyleSizePreviewRow = addRadioWithPreview(attributePanel, L("fontFamilyStyleSize"), initialAttributePreview.familyStyleSize, ATTRIBUTE_LABEL_WIDTH);
    var fontSizePreviewRow = addRadioWithPreview(attributePanel, L("fontSize"), initialAttributePreview.size, ATTRIBUTE_LABEL_WIDTH);

    var rbFontFamily = fontFamilyPreviewRow.radioButton;
    var rbFontFamilyStyle = fontFamilyStylePreviewRow.radioButton;
    var rbFontFamilyStyleSize = fontFamilyStyleSizePreviewRow.radioButton;
    var rbFontSize = fontSizePreviewRow.radioButton;
    var rbTextFillColor = addRadioOnlyRow(attributePanel, L("textFillColor"), ATTRIBUTE_LABEL_WIDTH).radioButton;
    var opacityPreviewRow = addRadioWithPreview(attributePanel, L("opacity"), initialAttributePreview.opacity, ATTRIBUTE_LABEL_WIDTH);
    var rbOpacity = opacityPreviewRow.radioButton;

    setHelpTip(rbFontFamily, L("tipFontFamily"));
    setHelpTip(rbFontFamilyStyle, L("tipFontFamilyStyle"));
    setHelpTip(rbFontFamilyStyleSize, L("tipFontFamilyStyleSize"));
    setHelpTip(rbFontSize, L("tipFontSize"));
    setHelpTip(rbTextFillColor, L("tipTextFillColor"));
    setHelpTip(rbOpacity, L("tipOpacity"));

    /* テキスト種類と文字列条件を横並びに配置 / Arrange text type and string condition panels side by side */
    var textConditionGroup = selectionPanel.add("group");
    textConditionGroup.orientation = "row";
    textConditionGroup.alignChildren = "fill";
    textConditionGroup.alignment = "fill";

    /* テキスト種類パネルを追加 / Add text type panel */
    var textTypePanel = textConditionGroup.add("panel", undefined, L("textTypePanel"));
    setupPanelLayout(textTypePanel, 6);

    /* テキスト種類ラジオボタン / Text type radio buttons */
    var rbAllText = textTypePanel.add("radiobutton", undefined, L("allText"));
    var rbPointText = textTypePanel.add("radiobutton", undefined, L("pointText"));
    var rbAreaText = textTypePanel.add("radiobutton", undefined, L("areaText"));
    var rbPathText = textTypePanel.add("radiobutton", undefined, L("pathText"));

    setHelpTip(rbAllText, L("tipAllText"));
    setHelpTip(rbPointText, L("tipPointText"));
    setHelpTip(rbAreaText, L("tipAreaText"));
    setHelpTip(rbPathText, L("tipPathText"));

    /* 初期選択を設定 / Set initial selection */
    rbAllText.value = true;

    /* 文字列条件パネルを作成 / Create string condition panel */
    var textMatchPanelTitle = L("textMatchPanel");
    var textMatchPanel = textConditionGroup.add("panel", undefined, textMatchPanelTitle);
    setupPanelLayout(textMatchPanel, 6);

    var textKeywordInput = textMatchPanel.add("edittext", undefined, initialString);
    textKeywordInput.characters = 22;
    setHelpTip(textKeywordInput, L("tipKeywordInput"));

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

    var rbExactMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, L("exactMatch"));
    var rbLooseMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, L("containsMatch"));

    var rbStartsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, L("startsWith"));
    var rbEndsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, L("endsWith"));

    var rbRegexMatch = textMatchRightColumnGroup.add("radiobutton", undefined, L("regexMatch"));

    setHelpTip(rbExactMatch, L("tipExactMatch"));
    setHelpTip(rbStartsWith, L("tipStartsWith"));
    setHelpTip(rbEndsWith, L("tipEndsWith"));
    setHelpTip(rbLooseMatch, L("tipContainsMatch"));
    setHelpTip(rbRegexMatch, L("tipRegexMatch"));

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
            alert(L("errEmptyKeyword"));
            return null;
        }
        if (matchMode === "regex") {
            try {
                return { regex: new RegExp(keyword) };
            } catch (regexError) {
                alert(L("errInvalidRegex"));
                return null;
            }
        }
        return { regex: null };
    }

    /* 文字列条件でテキストを選択 / Select text frames by string condition */
    function selectTextFramesByString(keyword, matchMode, regex) {
        var doc = app.activeDocument;
        var allItems = doc.textFrames;
        var selected = [];

        for (var i = 0; i < allItems.length; i++) {
            var textFrame = allItems[i];
            var textContents = textFrame.contents || "";

            if (textMatches(textContents, keyword, matchMode, regex)) {
                selected.push(textFrame);
            }
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
    var optionPanel = dialog.add("panel", undefined, L("postProcessPanel"));
    setupPanelLayout(optionPanel, 6);
    /* ドキュメント内に TextFrame が 0 件なら後処理は無意味なのでディム / Disable post-process when document has no text frames */
    optionPanel.enabled = app.activeDocument.textFrames.length > 0;

    var rbNoPostProcess = optionPanel.add("radiobutton", undefined, L("noPostProcess"));
    var rbHide = optionPanel.add("radiobutton", undefined, L("hideAfterSelection"));
    var rbMove = optionPanel.add("radiobutton", undefined, L("moveToTextLayer"));
    var rbBulkEdit = optionPanel.add("radiobutton", undefined, L("bulkEdit"));

    rbNoPostProcess.value = true;

    setHelpTip(rbNoPostProcess, L("tipNoPostProcess"));
    setHelpTip(rbHide, L("tipHideAfterSelection"));
    setHelpTip(rbMove, L("tipMoveToTextLayer"));
    setHelpTip(rbBulkEdit, L("tipBulkEdit"));

    setupExclusiveRadioButtons([
        rbNoPostProcess,
        rbHide,
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

        var bulkDialog = new Window("dialog", L("bulkEditDialogTitle"));
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
        var bulkCancelBtn = bulkButtonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
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
            var layerInfo = getOrCreateLayerByName(doc, "_text");
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
    // テキスト選択処理 / Text selection processing
    // =========================================
    /* 種類ごとにテキストを選択 / Select text frames by text type */
    function selectTextFramesByType(textType) {

        var doc = app.activeDocument;
        var allItems = doc.textFrames;
        var selected = [];

        for (var i = 0; i < allItems.length; i++) {
            var textFrame = allItems[i];

            if (textType === "all") {
                selected.push(textFrame);

            } else if (textType === "point") {
                /* ポイント文字を判定 / Check point text */
                if (textFrame.kind === TextType.POINTTEXT) {
                    selected.push(textFrame);
                }

            } else if (textType === "area") {
                /* エリア内文字を判定 / Check area text */
                if (textFrame.kind === TextType.AREATEXT) {
                    selected.push(textFrame);
                }

            } else if (textType === "path") {
                /* パス上文字を判定 / Check path text */
                if (textFrame.kind === TextType.PATHTEXT) {
                    selected.push(textFrame);
                }
            }
        }

        /* 選択を適用 / Apply selection */
        doc.selection = selected;
        return selected.length;
    }

    // =========================================
    // ボタンエリア / Button area
    // =========================================
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelBtn = buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    /* OKボタン実行処理 / Handle OK button action */
    okBtn.onClick = function () {
        var selectedPostProcessMode = getSelectedPostProcessMode();
        var selectedAttributeCommand = getSelectedAttributeCommand();
        if (selectedAttributeCommand) {
            dialog.close();
            executeSelectedAttributeCommand(selectedAttributeCommand);
            applyPostProcessToSelection(selectedPostProcessMode);
            return;
        }

        var textMatchMode = getSelectedTextMatchMode();
        if (textMatchMode) {
            var keyword = textKeywordInput.text || "";
            var validation = validateTextMatchInput(keyword, textMatchMode);
            if (!validation) {
                return;
            }
            dialog.close();
            var matchedCount = selectTextFramesByString(keyword, textMatchMode, validation.regex);
            if (matchedCount === 0) {
                alert(L("errNoMatch"));
                return;
            }
            applyPostProcessToSelection(selectedPostProcessMode);
            return;
        }

        var selectedTextType = getSelectedTextType();
        dialog.close();
        var typeMatchedCount = selectTextFramesByType(selectedTextType);
        if (typeMatchedCount === 0) {
            alert(L("errNoMatch"));
            return;
        }
        applyPostProcessToSelection(selectedPostProcessMode);
    };

    /* キャンセルボタン処理 / Handle Cancel button action */
    cancelBtn.onClick = function () {
        dialog.close();
    };

    /* ダイアログを表示 / Show dialog */
    dialog.show();
}());