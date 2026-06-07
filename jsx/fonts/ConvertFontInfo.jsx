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
- Enter／Returnキーで確定（OKがデフォルトボタン）
- 詳細表示ではラベル行と値行を分け、行揃えを左揃えに変更
- 選択したテキストフレームを保持し、プレビュー中に選択状態が変わっても同じ対象を更新
- フォントサイズは環境設定「文字の単位」に従い、小数第2位まで換算表示
- 日本語／英語インターフェース対応


### 更新履歴

- v1.0.0 (20250509) : 初期バージョン
- v1.1.0 (20260428) : 変換形式を拡張し、実フォント情報プレビュー、キー操作、OK時の再適用を追加
- v1.1.1 (20260608) : 文字の部分選択（TextRange）からも実行可能に、OKをデフォルトボタン化、内部リファクタ

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
- Confirm with the Enter / Return key (OK is the default button)
- Detail view separates label lines and value lines, and sets justification to left
- Keeps references to the selected text frames, so the same targets are updated even if the selection changes during preview
- Converts font size according to the Illustrator "Type Units" preference and rounds it to two decimals
- Japanese and English UI support


### Update History

- v1.0.0 (20250509): Initial version
- v1.1.0 (20260428): Expanded conversion formats and added actual font preview, keyboard shortcuts, and final reapply on OK
*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.1";

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var LABELS = {
        dialog: {
            title: { ja: "フォント情報に変換", en: "Convert Font Info" }
        },
        panel: {
            title: { ja: "変換形式", en: "Conversion Format" }
        },
        radio: {
            family: { ja: "フォント名", en: "Font Family" },
            style: { ja: "スタイル", en: "Style" },
            familyStyle: { ja: "フォント名＋スタイル", en: "Font Family + Style" },
            postscript: { ja: "PostScript 名", en: "PostScript Name" },
            fullNameSize: { ja: "フルネーム＋サイズ", en: "Full Name + Size" },
            detail: { ja: "詳細", en: "Labeled Detail" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        help: {
            fontFamily: { ja: "ショートカット: F", en: "Shortcut: F" },
            style: { ja: "ショートカット: S", en: "Shortcut: S" },
            fontFamilyStyle: { ja: "ショートカット: B", en: "Shortcut: B" },
            postScriptName: { ja: "ショートカット: P", en: "Shortcut: P" },
            fullNameSize: { ja: "ショートカット: M", en: "Shortcut: M" },
            detailLines: { ja: "ショートカット: D", en: "Shortcut: D" }
        },
        detail: {
            familyLabel: { ja: "フォント名", en: "Font Family" },
            styleLabel: { ja: "スタイル", en: "Style" },
            nameLabel: { ja: "PostScript 名", en: "PostScript Name" }
        }
    };

    var currentLanguage = getCurrentLang();

    function L(key) {
        var labelEntry = LABELS;
        var keyParts = key.split('.');
        for (var keyPartIndex = 0; keyPartIndex < keyParts.length; keyPartIndex++) {
            labelEntry = labelEntry && labelEntry[keyParts[keyPartIndex]];
        }
        if (!labelEntry) return key;
        return labelEntry[currentLanguage] || labelEntry.ja || labelEntry.en || key;
    }

    // =========================================
    // 単位 / Units
    // =========================================
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

    // =========================================
    // UI設定 / UI Settings
    // =========================================
    var FORMAT_ITEMS = [
        { labelKey: 'radio.family', format: 'name', helpKey: 'help.fontFamily', shortcut: 'F', showPreview: true },
        { labelKey: 'radio.style', format: 'style', helpKey: 'help.style', shortcut: 'S', showPreview: true },
        { labelKey: 'radio.familyStyle', format: 'family+style', helpKey: 'help.fontFamilyStyle', shortcut: 'B', showPreview: true },
        { labelKey: 'radio.postscript', format: 'postscript', helpKey: 'help.postScriptName', shortcut: 'P', showPreview: true },
        { labelKey: 'radio.fullNameSize', format: 'fullName+size', helpKey: 'help.fullNameSize', shortcut: 'M', showPreview: true },
        { labelKey: 'radio.detail', format: 'detailLines', helpKey: 'help.detailLines', shortcut: 'D', showPreview: false }
    ];

    // 文字ツールで文字を選択すると、app.selection は配列ではなく TextRange を返す
    // （配列でなく .story を持つオブジェクトを TextRange とみなす）
    function isTextRangeSelection(selection) {
        return selection && !(selection instanceof Array) && selection.story;
    }

    // TextRange（文字選択）から、その文字が属する親テキストフレームを取得
    // contents はストーリー単位なので、連結テキストでも先頭フレーム1つで足りる
    function textFramesFromTextRange(textRange) {
        var storyTextFrames = textRange.story && textRange.story.textFrames;
        return (storyTextFrames && storyTextFrames.length >= 1) ? [storyTextFrames[0]] : [];
    }

    // 選択内容から対象テキストフレームの配列を解決する（doc.selection は書き換えない）
    // ・オブジェクト選択：選択内のテキストフレーム
    // ・文字ツールでの部分選択（TextRange）：その文字の親テキストフレーム
    function resolveTargetTextFrames(doc) {
        var selection = doc.selection;
        if (!selection) return [];

        if (isTextRangeSelection(selection)) {
            return textFramesFromTextRange(selection);
        }
        if (!(selection instanceof Array)) return [];

        var textFrames = [];
        for (var selectedIndex = 0; selectedIndex < selection.length; selectedIndex++) {
            var selectedItem = selection[selectedIndex];
            if (selectedItem.typename === "TextFrame") {
                textFrames.push(selectedItem);
            } else if (isTextRangeSelection(selectedItem)) {
                // 配列内に TextRange が入るケースにも対応
                var parentFrames = textFramesFromTextRange(selectedItem);
                for (var parentIndex = 0; parentIndex < parentFrames.length; parentIndex++) {
                    textFrames.push(parentFrames[parentIndex]);
                }
            }
        }
        return textFrames;
    }

    function collectTextFrameInfo(textFrame) {
        try {
            var textRange = textFrame.textRange;
            return {
                textFrame: textFrame,
                font: textRange.characterAttributes.textFont,
                originalSize: textRange.characterAttributes.size,
                originalJustification: textRange.paragraphAttributes.justification,
                originalText: textFrame.contents
            };
        } catch (e) {
            return null;
        }
    }

    function collectSelectedTextFrameInfos(selectedItems) {
        var textFrameInfos = [];

        // 配列風オブジェクト以外（TextRange など）は対象外
        if (!selectedItems || typeof selectedItems.length !== 'number') {
            return textFrameInfos;
        }

        for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
            var selectedItem = selectedItems[selectedIndex];
            if (selectedItem.typename !== "TextFrame") continue;

            var textFrameInfo = collectTextFrameInfo(selectedItem);
            if (textFrameInfo) {
                textFrameInfos.push(textFrameInfo);
            }
        }

        return textFrameInfos;
    }

    function restoreTextFrameInfo(originalTextFrameInfo) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            textFrame.contents = originalTextFrameInfo.originalText;
            var textRange = textFrame.textRange;
            textRange.characterAttributes.textFont = originalTextFrameInfo.font;
            textRange.characterAttributes.size = originalTextFrameInfo.originalSize;
            textRange.paragraphAttributes.justification = originalTextFrameInfo.originalJustification;
        } catch (e) {
            return false;
        }
        return true;
    }

    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var targetTextFrames = resolveTargetTextFrames(doc);
        if (targetTextFrames.length === 0 || targetTextFrames.length >= 1000) return;

        var originalTextFrameInfos = collectSelectedTextFrameInfos(targetTextFrames);
        if (originalTextFrameInfos.length === 0) return;

        showDialog(originalTextFrameInfos);
    }

    function getConvertedFontInfoText(originalTextFrameInfo, format) {
        var sourceFont = originalTextFrameInfo.font;
        var sourceFontSize = originalTextFrameInfo.originalSize;

        if (!sourceFont) return "";

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
            { label: L('detail.familyLabel'), value: sourceFont.family },
            { label: L('detail.styleLabel'), value: sourceFont.style },
            { label: L('detail.nameLabel'), value: sourceFont.name }
        ];
    }

    function findTextFontByName(fontName) {
        var textFonts = app.textFonts;
        for (var fontIndex = 0; fontIndex < textFonts.length; fontIndex++) {
            if (textFonts[fontIndex].name === fontName) {
                return textFonts[fontIndex];
            }
        }
        return null;
    }

    // 詳細表示のラベル用フォント（セッション中変わらないので一度だけ走査してキャッシュ）
    var cachedDetailLabelFont; // 未取得は undefined、取得済みで未発見は null
    function getDetailLabelFont() {
        if (cachedDetailLabelFont === undefined) {
            cachedDetailLabelFont = findTextFontByName("HiraginoSans-W3");
        }
        return cachedDetailLabelFont;
    }

    function buildDetailLinesText(originalTextFrameInfo) {
        var lineBreak = String.fromCharCode(13);
        var detailLineItems = buildDetailLineItems(originalTextFrameInfo);
        var detailTextParts = [];

        for (var detailItemIndex = 0; detailItemIndex < detailLineItems.length; detailItemIndex++) {
            detailTextParts.push(detailLineItems[detailItemIndex].label);
            detailTextParts.push(detailLineItems[detailItemIndex].value);
        }

        return detailTextParts.join(lineBreak);
    }

    function applyDetailFontInfoPreview(originalTextFrameInfo, detailLabelFont) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;

            textFrame.contents = buildDetailLinesText(originalTextFrameInfo);
            textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;

            var previewLines = textFrame.textRange.lines;
            if (previewLines.length > 0 && detailLabelFont) {
                for (var lineIndex = 0; lineIndex < previewLines.length; lineIndex++) {
                    var isLabelLine = (lineIndex % 2 === 0);
                    var lineAttributes = previewLines[lineIndex].characterAttributes;
                    lineAttributes.textFont = isLabelLine ? detailLabelFont : sourceFont;
                    lineAttributes.size = isLabelLine ? 10 : sourceFontSize;
                }
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    function applySimpleFontInfoPreview(originalTextFrameInfo, format) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;
            var textRange = textFrame.textRange;

            textRange.characterAttributes.textFont = sourceFont;
            textRange.characterAttributes.size = sourceFontSize;
            textFrame.contents = getConvertedFontInfoText(originalTextFrameInfo, format);
        } catch (e) {
            return false;
        }
        return true;
    }

    function updateFontInfoPreview(originalTextFrameInfos, format) {
        var isDetailLinesFormat = (format === "detailLines");
        var detailLabelFont = isDetailLinesFormat ? getDetailLabelFont() : null;

        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            if (isDetailLinesFormat) {
                applyDetailFontInfoPreview(originalTextFrameInfos[textFrameIndex], detailLabelFont);
            } else {
                applySimpleFontInfoPreview(originalTextFrameInfos[textFrameIndex], format);
            }
        }

        app.redraw();
    }

    function restoreOriginalText(originalTextFrameInfos) {
        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            restoreTextFrameInfo(originalTextFrameInfos[textFrameIndex]);
        }

        app.redraw();
    }

    function showDialog(originalTextFrameInfos) {
        var dialog = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);

        var dialogContainer = dialog.add("group");
        dialogContainer.orientation = "column";
        dialogContainer.alignChildren = "left";
        dialogContainer.alignment = "fill";

        var getSelectedFormat = buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos);
        buildButtonRow(dialog, dialogContainer);

        var dialogResult = dialog.show();
        if (dialogResult === 1) {
            updateFontInfoPreview(originalTextFrameInfos, getSelectedFormat());
        } else {
            restoreOriginalText(originalTextFrameInfos);
        }
    }

    // 変換形式パネル（ラジオ生成・選択・キー操作）を組み立て、選択中フォーマットを返す関数を返す
    function buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos) {
        var DEFAULT_FORMAT = 'family+style';
        var previewSourceInfo = originalTextFrameInfos[0];
        var radioColumnWidth = 180;
        var formatRadioButtons = [];

        var formatPanel = dialogContainer.add("panel", undefined, L('panel.title'));
        formatPanel.orientation = "column";
        formatPanel.alignChildren = "left";
        formatPanel.alignment = "fill";
        formatPanel.margins = [15, 20, 15, 10];

        var formatRowsGroup = formatPanel.add("group");
        formatRowsGroup.orientation = "column";
        formatRowsGroup.alignChildren = ["fill", "top"];
        formatRowsGroup.spacing = 6;

        function addFormatRow(labelKey, previewText, helpTipKey) {
            var formatRow = formatRowsGroup.add("group");
            formatRow.orientation = "row";
            formatRow.alignChildren = ["left", "center"];

            var radioColumn = formatRow.add("group");
            radioColumn.preferredSize.width = radioColumnWidth;
            var radioButton = radioColumn.add("radiobutton", undefined, L(labelKey));

            if (helpTipKey) {
                radioButton.helpTip = L(helpTipKey);
            }

            if (previewText) {
                var previewStaticText = formatRow.add("statictext", undefined, previewText);
                previewStaticText.helpTip = previewText;
            }

            return radioButton;
        }

        function selectFormat(format) {
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                formatRadioButtons[itemIndex].value = (FORMAT_ITEMS[itemIndex].format === format);
            }
            updateFontInfoPreview(originalTextFrameInfos, format);
        }

        for (var formatIndex = 0; formatIndex < FORMAT_ITEMS.length; formatIndex++) {
            (function (formatItem, itemIndex) {
                var previewText = formatItem.showPreview ? getConvertedFontInfoText(previewSourceInfo, formatItem.format) : "";
                var radioButton = addFormatRow(formatItem.labelKey, previewText, formatItem.helpKey);
                formatRadioButtons[itemIndex] = radioButton;
                radioButton.onClick = function () {
                    selectFormat(formatItem.format);
                };
            })(FORMAT_ITEMS[formatIndex], formatIndex);
        }

        dialog.addEventListener("keydown", function (event) {
            var keyName = String(event.keyName).toUpperCase();
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                if (FORMAT_ITEMS[itemIndex].shortcut === keyName) {
                    selectFormat(FORMAT_ITEMS[itemIndex].format);
                    event.preventDefault();
                    return;
                }
            }
        });

        selectFormat(DEFAULT_FORMAT);

        return function () {
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                if (formatRadioButtons[itemIndex].value) return FORMAT_ITEMS[itemIndex].format;
            }
            return DEFAULT_FORMAT;
        };
    }

    // キャンセル／OK ボタン行を組み立て、OK をデフォルトボタンに設定
    function buildButtonRow(dialog, dialogContainer) {
        var buttonGroup = dialogContainer.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["fill", "center"];
        buttonGroup.alignment = "fill";
        buttonGroup.margins = [0, 10, 0, 0];

        var cancelButtonGroup = buttonGroup.add("group");
        cancelButtonGroup.orientation = "row";
        cancelButtonGroup.alignChildren = ["left", "center"];
        cancelButtonGroup.alignment = ["left", "center"];
        cancelButtonGroup.add("button", undefined, L('button.cancel'), { name: "cancel" });

        var centerSpacerGroup = buttonGroup.add("group");
        centerSpacerGroup.alignment = ["fill", "fill"];
        centerSpacerGroup.minimumSize.width = 100;

        var okButtonGroup = buttonGroup.add("group");
        okButtonGroup.orientation = "row";
        okButtonGroup.alignChildren = ["right", "center"];
        okButtonGroup.alignment = ["right", "center"];
        var okButton = okButtonGroup.add("button", undefined, L('button.ok'), { name: "ok" });
        dialog.defaultElement = okButton;
    }

    main();

})();