#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ConvertFontInfo.jsx

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ConvertFontInfo.jsx


### 概要

- 選択したテキストをフォント情報に変換するIllustrator用スクリプトです。
- ダイアログで変換形式を選び、対象テキストを即時プレビューしながら書き換えます。
- 右列には、選択中テキストの実フォント情報をもとにした変換結果を表示します。
- OK時は選択中の形式を再適用し、キャンセル時は開始時の内容、フォント、サイズ、行揃えに戻します。

### 主な機能

- フォント名、スタイル、フォント名＋スタイル、PostScript名、フルネーム＋サイズ、詳細表示を選択可能
- 初期状態は「フォント名＋スタイル」を選択
- F/S/B/P/M/Dキーで変換形式を切り替え
- 詳細表示ではラベル行と値行を分け、行揃えを左揃えに変更
- 選択したテキストフレームを保持し、プレビュー中に選択状態が変わっても同じ対象を更新
- フォントサイズは環境設定「文字の単位」に従い、小数第2位まで換算表示
- 日本語／英語インターフェース対応


### 更新履歴

- v1.0.0 (20250509) : 初期バージョン
- v1.1.0 (20260428) : 変換形式を拡張し、実フォント情報プレビュー、キー操作、OK時の再適用を追加

---

### Script Name:

ConvertFontInfo.jsx

### Overview

- An Illustrator script that converts selected text into font information.
- Choose a conversion format in the dialog and rewrite the target text with an immediate preview.
- The right column shows conversion results generated from the actual font information of the selected text.
- Reapplies the selected format on OK, and restores original contents, font, size, and justification when canceled.

### Main Features

- Choose from font family, style, font family + style, PostScript name, full name + size, or detail view
- Defaults to Font Family + Style
- Switch conversion formats with the F/S/B/P/M/D keys
- Detail view separates label lines and value lines, and sets justification to left
- Keeps references to the selected text frames, so the same targets are updated even if the selection changes during preview
- Converts font size according to the Illustrator "Type Units" preference and rounds it to two decimals
- Japanese and English UI support


### Update History

- v1.0.0 (20250509): Initial version
- v1.1.0 (20260428): Expanded conversion formats and added actual font preview, keyboard shortcuts, and final reapply on OK
*/

(function () {

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();

    var SCRIPT_VERSION = "v1.1.0";

    function L(key) {
        if (!LABELS[key]) {
            $.writeln("[ConvertFontInfo] Missing LABELS key: " + key);
            return "[" + key + "]";
        }

        if (LABELS[key][lang]) {
            return LABELS[key][lang];
        }

        if (LABELS[key].en) {
            $.writeln("[ConvertFontInfo] Missing locale '" + lang + "' for LABELS key: " + key + ". Fallback to en.");
            return LABELS[key].en;
        }

        $.writeln("[ConvertFontInfo] Missing locale '" + lang + "' and fallback 'en' for LABELS key: " + key);
        return "[" + key + "]";
    }

    var LABELS = {
        dialogTitle: { ja: "フォント情報に変換", en: "Convert Font Info" },
        panelTitle: { ja: "変換形式", en: "Conversion Format" },
        radio1: { ja: "フォント名", en: "Font Family" },
        radio2: { ja: "スタイル", en: "Style" },
        radio3: { ja: "フォント名＋スタイル", en: "Font Family + Style" },
        radio4: { ja: "PostScript 名", en: "PostScript Name" },
        radio5: { ja: "フルネーム＋サイズ", en: "Full Name + Size" },
        radio6: { ja: "詳細", en: "Labeled Detail" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" },
        helpFontFamily: { ja: "ショートカット: F", en: "Shortcut: F" },
        helpStyle: { ja: "ショートカット: S", en: "Shortcut: S" },
        helpFontFamilyStyle: { ja: "ショートカット: B", en: "Shortcut: B" },
        helpPostScriptName: { ja: "ショートカット: P", en: "Shortcut: P" },
        helpFullNameSize: { ja: "ショートカット: M", en: "Shortcut: M" },
        helpDetailLines: { ja: "ショートカット: D", en: "Shortcut: D" },
        detailFamilyLabel: { ja: "フォント名", en: "Font Family" },
        detailStyleLabel: { ja: "スタイル", en: "Style" },
        detailNameLabel: { ja: "PostScript 名", en: "PostScript Name" }
    };

    // 単位変換処理（pt → 指定単位、小数第2位）
    function getFontSizeWithUnit(sizePt) {
        var unitPref = app.preferences.getIntegerPreference('text/units');
        var unitLabel = "pt";
        var convertedSize = sizePt;

        switch (unitPref) {
            case 0: unitLabel = "inch"; convertedSize = sizePt / 72; break;
            case 1: unitLabel = "mm"; convertedSize = sizePt * 25.4 / 72; break;
            case 2: unitLabel = "pt"; convertedSize = sizePt; break;
            case 3: unitLabel = "pica"; convertedSize = sizePt / 12; break;
            case 4: unitLabel = "cm"; convertedSize = sizePt * 2.54 / 72; break;
            case 5: unitLabel = "Q"; convertedSize = sizePt * (25.4 / 72) * 4; break;
            case 6: unitLabel = "px"; convertedSize = sizePt * (96 / 72); break;
        }

        return (Math.round(convertedSize * 100) / 100) + " " + unitLabel;
    }

    var originalTextFrameInfos = [];

    main();

    function main() {
        if (app.documents.length === 0) return;

        if (app.selection.constructor.name === "TextRange") {
            var frames = app.selection.story.textFrames;
            if (frames.length === 1) app.selection = [frames[0]];
        }

        var selectedItems = app.selection;
        if (!selectedItems || selectedItems.length === 0 || selectedItems.length >= 1000) return;

        for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
            var selectedItem = selectedItems[selectedIndex];
            if (selectedItem.typename !== "TextFrame") continue;
            try {
                var textRange = selectedItem.textRange;
                originalTextFrameInfos.push({
                    textFrame: selectedItem,
                    font: textRange.characterAttributes.textFont,
                    originalSize: textRange.characterAttributes.size,
                    originalJustification: textRange.paragraphAttributes.justification,
                    originalText: selectedItem.contents
                });
            } catch (e) { }
        }

        if (originalTextFrameInfos.length === 0) return;

        showDialog();
    }

    function getConvertedFontInfoText(originalTextFrameInfo, format) {
        var sourceFont = originalTextFrameInfo.font;
        var sourceFontSize = originalTextFrameInfo.originalSize;

        switch (format) {
            case "style":
                return sourceFont.style;
            case "family+style":
                return sourceFont.family + " " + sourceFont.style;
            case "postscript":
                return sourceFont.name;
            case "fullName+size":
                var fontSizeText = getFontSizeWithUnit(sourceFontSize);
                return (sourceFont.fullName && sourceFont.fullName !== "")
                    ? sourceFont.fullName + "\t" + fontSizeText
                    : sourceFont.family + " " + sourceFont.style + "\t" + fontSizeText;
            default:
                return sourceFont.family;
        }
    }

    function buildDetailLineItems(originalTextFrameInfo) {
        var sourceFont = originalTextFrameInfo.font;

        return [
            { label: L('detailFamilyLabel'), value: sourceFont.family },
            { label: L('detailStyleLabel'), value: sourceFont.style },
            { label: L('detailNameLabel'), value: sourceFont.name }
        ];
    }

    function updateFontInfoPreview(format) {
        var detailLabelFont = null;

        try {
            detailLabelFont = textFonts.getByName("HiraginoSans-W3");
        } catch (e) { }

        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            var originalTextFrameInfo = originalTextFrameInfos[textFrameIndex];
            var textFrame = originalTextFrameInfo.textFrame;

            try {
                var sourceFont = originalTextFrameInfo.font;
                var sourceFontSize = originalTextFrameInfo.originalSize;
                var convertedText = "";

                if (format === "label3lines") {
                    var lineBreak = String.fromCharCode(13);
                    var detailLineItems = buildDetailLineItems(originalTextFrameInfo);
                    var detailTextParts = [];

                    for (var detailItemIndex = 0; detailItemIndex < detailLineItems.length; detailItemIndex++) {
                        detailTextParts.push(detailLineItems[detailItemIndex].label);
                        detailTextParts.push(detailLineItems[detailItemIndex].value);
                    }

                    convertedText = detailTextParts.join(lineBreak);
                    textFrame.contents = convertedText;
                    textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;

                    var previewLines = textFrame.textRange.lines;
                    if (previewLines.length > 0 && detailLabelFont) {
                        for (var lineIndex = 0; lineIndex < previewLines.length; lineIndex++) {
                            var isLabelLine = (lineIndex % 2 === 0);
                            var attributes = previewLines[lineIndex].characterAttributes;
                            attributes.textFont = isLabelLine ? detailLabelFont : sourceFont;
                            attributes.size = isLabelLine ? 10 : sourceFontSize;
                        }
                    }

                } else {
                    var textRange = textFrame.textRange;
                    textRange.characterAttributes.textFont = sourceFont;
                    textRange.characterAttributes.size = sourceFontSize;

                    convertedText = getConvertedFontInfoText(originalTextFrameInfo, format);

                    textFrame.contents = convertedText;
                }

            } catch (e) { }
        }

        app.redraw();
    }

    function restoreOriginalText() {
        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            var originalTextFrameInfo = originalTextFrameInfos[textFrameIndex];
            var textFrame = originalTextFrameInfo.textFrame;

            try {
                textFrame.contents = originalTextFrameInfo.originalText;
                var textRange = textFrame.textRange;
                textRange.characterAttributes.textFont = originalTextFrameInfo.font;
                textRange.characterAttributes.size = originalTextFrameInfo.originalSize;
                textRange.paragraphAttributes.justification = originalTextFrameInfo.originalJustification;
            } catch (e) { }
        }

        app.redraw();
    }

    function showDialog() {
        var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);

        var dialogContainer = dialog.add("group");
        dialogContainer.orientation = "column";
        dialogContainer.alignChildren = "left";
        dialogContainer.alignment = "fill";
        var radioPanel = dialogContainer.add("panel", undefined, L('panelTitle'));
        radioPanel.orientation = "column";
        radioPanel.alignChildren = "left";
        radioPanel.alignment = "fill";
        radioPanel.margins = [15, 20, 15, 10];

        var formatRowsGroup = radioPanel.add("group");
        formatRowsGroup.orientation = "column";
        formatRowsGroup.alignChildren = ["fill", "top"];
        formatRowsGroup.spacing = 6;

        var previewSourceInfo = originalTextFrameInfos[0];
        var leftColumnWidth = 180;

        function addFormatRow(labelKey, previewText, helpTipKey) {
            var formatRow = formatRowsGroup.add("group");
            formatRow.orientation = "row";
            formatRow.alignChildren = ["left", "center"];

            var leftColumn = formatRow.add("group");
            leftColumn.preferredSize.width = leftColumnWidth;
            var radioButton = leftColumn.add("radiobutton", undefined, L(labelKey));

            if (helpTipKey) {
                radioButton.helpTip = L(helpTipKey);
            }

            if (previewText) {
                var previewStaticText = formatRow.add("statictext", undefined, previewText);
                previewStaticText.helpTip = previewText;
            }

            return radioButton;
        }

        var fontFamilyRadio = addFormatRow('radio1', getConvertedFontInfoText(previewSourceInfo, "name"), 'helpFontFamily');
        var styleRadio = addFormatRow('radio2', getConvertedFontInfoText(previewSourceInfo, "style"), 'helpStyle');
        var fontFamilyStyleRadio = addFormatRow('radio3', getConvertedFontInfoText(previewSourceInfo, "family+style"), 'helpFontFamilyStyle');
        var postScriptNameRadio = addFormatRow('radio4', getConvertedFontInfoText(previewSourceInfo, "postscript"), 'helpPostScriptName');
        var fullNameSizeRadio = addFormatRow('radio5', getConvertedFontInfoText(previewSourceInfo, "fullName+size"), 'helpFullNameSize');
        var detailLinesRadio = addFormatRow('radio6', "", 'helpDetailLines');

        fontFamilyStyleRadio.value = true;

        function setExclusiveRadio(selectedRadio) {
            fontFamilyRadio.value = false;
            styleRadio.value = false;
            fontFamilyStyleRadio.value = false;
            postScriptNameRadio.value = false;
            fullNameSizeRadio.value = false;
            detailLinesRadio.value = false;
            selectedRadio.value = true;
        }

        function selectFormatRadio(radioButton, format) {
            setExclusiveRadio(radioButton);
            updateFontInfoPreview(format);
        }

        function bindFormatRadio(radioButton, format) {
            radioButton.onClick = function () {
                selectFormatRadio(radioButton, format);
            };
        }

        function addFormatKeyHandler(dialog) {
            dialog.addEventListener("keydown", function (event) {
                var keyName = String(event.keyName).toUpperCase();

                if (keyName === "F") {
                    selectFormatRadio(fontFamilyRadio, "name");
                    event.preventDefault();
                } else if (keyName === "S") {
                    selectFormatRadio(styleRadio, "style");
                    event.preventDefault();
                } else if (keyName === "B") {
                    selectFormatRadio(fontFamilyStyleRadio, "family+style");
                    event.preventDefault();
                } else if (keyName === "P") {
                    selectFormatRadio(postScriptNameRadio, "postscript");
                    event.preventDefault();
                } else if (keyName === "M") {
                    selectFormatRadio(fullNameSizeRadio, "fullName+size");
                    event.preventDefault();
                }
                else if (keyName === "D") {
                    selectFormatRadio(detailLinesRadio, "label3lines");
                    event.preventDefault();
                }
            });
        }

        bindFormatRadio(fontFamilyRadio, "name");
        bindFormatRadio(styleRadio, "style");
        bindFormatRadio(fontFamilyStyleRadio, "family+style");
        bindFormatRadio(postScriptNameRadio, "postscript");
        bindFormatRadio(fullNameSizeRadio, "fullName+size");
        bindFormatRadio(detailLinesRadio, "label3lines");

        function getSelectedFormat() {
            if (fontFamilyRadio.value) return "name";
            if (styleRadio.value) return "style";
            if (fontFamilyStyleRadio.value) return "family+style";
            if (postScriptNameRadio.value) return "postscript";
            if (fullNameSizeRadio.value) return "fullName+size";
            if (detailLinesRadio.value) return "label3lines";
            return "family+style";
        }

        addFormatKeyHandler(dialog);

        updateFontInfoPreview("family+style");

        var buttonGroup = dialogContainer.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["fill", "center"];
        buttonGroup.alignment = "fill";
        buttonGroup.margins = [0, 10, 0, 0];

        var leftButtonGroup = buttonGroup.add("group");
        leftButtonGroup.orientation = "row";
        leftButtonGroup.alignChildren = ["left", "center"];
        leftButtonGroup.alignment = ["left", "center"];
        leftButtonGroup.add("button", undefined, L('cancel'), { name: "cancel" });

        var centerSpacerGroup = buttonGroup.add("group");
        centerSpacerGroup.alignment = ["fill", "fill"];
        centerSpacerGroup.minimumSize.width = 100;

        var rightButtonGroup = buttonGroup.add("group");
        rightButtonGroup.orientation = "row";
        rightButtonGroup.alignChildren = ["right", "center"];
        rightButtonGroup.alignment = ["right", "center"];
        rightButtonGroup.add("button", undefined, L('ok'), { name: "ok" });

        var dialogResult = dialog.show();
        if (dialogResult === 1) {
            updateFontInfoPreview(getSelectedFormat());
        } else {
            restoreOriginalText();
        }
    }

})();