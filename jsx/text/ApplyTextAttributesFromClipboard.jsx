#target illustrator
#targetengine "FontClipboard"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

「CopyTextAttributesToClipboard.jsx」が永続エンジン "FontClipboard" の
$.global.FontClipboard に保存した文字属性・段落属性を、選択中のテキストへ適用します。
テキスト編集モードで文字列を部分選択している場合はその範囲に、複数のテキストオブジェクトを選択している場合はそれぞれに同じ属性を適用します。
保存された内容は、フォント・スタイル・フォントサイズ・行送り・自動行送り・
トラッキング・自動カーニング・自動カーニングの表示名・プロポーショナルメトリクス・
組み方向・組み方向の表示名・行揃え・行揃えの表示名です。
PostScript名は内部のフォント適用に使用しますが、ダイアログには表示しません。
ダイアログでは適用する項目をチェックボックスで選択できます。
初期状態ではフォントのみONにし、Optionクリックでクリックした項目以外をOFFにできます。
フォントのスタイルは、独立した適用項目ではなく、フォント名の下に補足情報として表示します。
自動カーニングのうち、メトリクス・オプティカル・0は Illustrator ExtendScript で取得できても直接適用できないため、
ダイナミックアクションを一時的に読み込んで適用します。
プロポーショナルメトリクスは通常の文字属性として読み書きします。
プレビューをONにすると、ダイアログを閉じずに一時適用し、OFFまたはキャンセルでスクリプト開始前の属性へ戻します。
行揃えは段落属性として扱い、プレビュー解除・キャンセル時には段落ごとの行揃えも復元します。
フォントサイズと行送りは Illustrator 内部値の pt で適用し、ダイアログ上では text/units の設定に合わせて表示します。
選択範囲内で属性が混在（例：文字ごとに異なるフォントやサイズ）している場合、退避・復元は代表値での上書きとなり、元の混在状態は再現されません。
Illustrator を再起動するとエンジンが破棄され、記憶した値は消えます。

Overview

This script applies text and paragraph attributes stored by "CopyTextAttributesToClipboard.jsx" in
$.global.FontClipboard on the persistent "FontClipboard" engine to the selected text.
If a text range is partially selected in text editing mode, the attributes are applied to that range;
when multiple text objects are selected, the same attributes are applied to each object.
The stored values include font, style, font size, leading, auto leading,
tracking, auto kerning, auto kerning label, proportional metrics,
text orientation, orientation label, alignment, and alignment label.
The PostScript name is used internally for font application but is not shown in the dialog.
The dialog lets you choose which attributes to apply using checkboxes.
By default, only Font is enabled; Option-click turns off all items except the clicked one.
The font style is shown as supplemental information under the font name, not as an independent apply item.
For auto kerning, Metrics, Optical, and 0 can be read by Illustrator ExtendScript but cannot be applied directly,
so this script temporarily loads dynamic actions to apply them.
Proportional metrics are read and written as a normal character attribute.
When Preview is enabled, the selected attributes are applied temporarily without closing the dialog,
and turning Preview off or canceling restores the attributes from before the script started.
Alignment is handled as a paragraph attribute, and paragraph-level alignment is also restored when Preview is turned off or the dialog is canceled.
Font size and leading are applied as Illustrator's internal point values, while the dialog displays them
according to the text/units preference.
If the selection contains mixed attributes (for example, different fonts or sizes per character), capture and restore
overwrite the range with a single representative value, so the original mixed state cannot be reproduced.
When Illustrator is restarted, the engine is destroyed and the stored values are cleared.

作成日 / Created: 2026-04-25
更新日 / Updated: 2026-04-25
*/


var SCRIPT_VERSION = "v1.0";

function getCurrentLocaleLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLocaleLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "文字属性を適用",
        en: "Apply Text Attributes"
    },
    previewCheckbox: {
        ja: "プレビュー",
        en: "Preview"
    },
    errorNoDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    errorNoCopiedAttributes: {
        ja: "先に「文字属性をコピー」スクリプトを実行してください。",
        en: "Run the Copy Text Attributes script first."
    },
    errorNoTextSelection: {
        ja: "テキストオブジェクトを選択するか、テキスト編集モードで文字列を選択してください。",
        en: "Select text objects, or select text in text editing mode."
    },
    errorEmptyTextRange: {
        ja: "文字のある範囲を選択してください。",
        en: "Select a text range that contains characters."
    },
    errorNoApplyItem: {
        ja: "適用する項目をひとつ以上選んでください。",
        en: "Select at least one item to apply."
    },
    errorFontNotFoundPrefix: {
        ja: "フォント \"",
        en: "Font \""
    },
    errorFontNotFoundSuffix: {
        ja: "\" が見つかりませんでした。\nこの環境にインストールされていない可能性があります。",
        en: "\" was not found.\nIt may not be installed in this environment."
    },
    fontLabel: {
        ja: "フォント",
        en: "Font"
    },
    sizeLabel: {
        ja: "フォントサイズ",
        en: "Font size"
    },
    leadingLabel: {
        ja: "行送り",
        en: "Leading"
    },
    autoLeadingLabel: {
        ja: "自動行送り",
        en: "Auto leading"
    },
    trackingLabel: {
        ja: "トラッキング",
        en: "Tracking"
    },
    kerningMethodLabel: {
        ja: "自動カーニング",
        en: "Auto kerning"
    },
    proportionalMetricsLabel: {
        ja: "プロポーショナルメトリクス",
        en: "Proportional metrics"
    },
    orientationLabel: {
        ja: "組み方向",
        en: "Orientation"
    },
    justificationLabel: {
        ja: "行揃え",
        en: "Alignment"
    },
    horizontalOrientation: {
        ja: "横組み",
        en: "Horizontal"
    },
    verticalOrientation: {
        ja: "縦組み",
        en: "Vertical"
    },
    onValue: {
        ja: "ON",
        en: "On"
    },
    offValue: {
        ja: "OFF",
        en: "Off"
    },
    notStored: {
        ja: "（未記憶）",
        en: "(not stored)"
    },
    cancelButton: {
        ja: "キャンセル",
        en: "Cancel"
    },
    applyButton: {
        ja: "適用",
        en: "Apply"
    }
};

// =========================================
// UI設定 / UI settings
// =========================================

var ATTRIBUTE_LABEL_WIDTH = 185;

function L(key) {
    return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
}

function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}

// =========================================
// 単位ユーティリティ / Unit utilities
// =========================================

/* 単位コードとラベルのマップ / Unit code to label map */
var preferenceUnitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* 単位コードとpt換算係数のマップ / Unit code to point conversion factor map */
/* 1単位あたりのpt値を定義 / Defines point value per unit */
var preferenceUnitPointFactorMap = {
    0: 72,                 // in
    1: 72 / 25.4,          // mm
    2: 1,                  // pt
    3: 12,                 // pica
    4: 72 / 2.54,          // cm
    5: 72 / 25.4 * 0.25,   // Q/H（0.25mm） / Q/H (0.25 mm)
    6: 1,                  // px
    7: 72 * 12,            // ft/in
    8: 72 / 0.0254,        // m
    9: 72 * 36,            // yd
    10: 72 * 12            // ft
};

/* 単位コードと設定キーから適切な単位ラベルを返す / Return the proper unit label from unit code and preference key */
function getPreferenceUnitLabel(unitCode, prefKey) {
    if (unitCode === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return preferenceUnitLabelMap[unitCode] || "pt";
}

/* 環境設定から単位コードを取得 / Get unit code from preferences */
function getPreferenceUnitCode(prefKey, fallbackCode) {
    try {
        return app.preferences.getIntegerPreference(prefKey);
    } catch (e) {
        return fallbackCode;
    }
}

/* pt値を指定単位へ変換 / Convert point value to the specified unit */
function convertPointsToPreferenceUnit(pointValue, unitCode) {
    var pointFactor = preferenceUnitPointFactorMap[unitCode] || 1;
    return pointValue / pointFactor;
}

/* 数値を表示用に丸める / Round number for display */
function formatNumberForDisplay(value, digits) {
    var multiplier = Math.pow(10, digits);
    var rounded = Math.round(value * multiplier) / multiplier;
    return String(rounded);
}

/* pt値を現在の文字単位で表示 / Format point value using current text unit */
function formatPointValueForDialog(pointValue) {
    var unitCode = getPreferenceUnitCode("text/units", 2);
    var textUnitLabel = getPreferenceUnitLabel(unitCode, "text/units");
    var displayValue = convertPointsToPreferenceUnit(pointValue, unitCode);
    return formatNumberForDisplay(displayValue, 3) + " " + textUnitLabel;
}

// =========================================
// 文字属性ユーティリティ / Text attribute utilities
// =========================================

/* 選択から適用対象を取得 / Get apply targets from selection */
function getApplyTargetsFromSelection(selection) {
    var targets = [];

    if (selection && selection.typename === "TextRange") {
        targets.push({
            textRange: selection,
            textFrame: getParentTextFrameFromTextRange(selection)
        });
        return targets;
    }

    if (selection && selection.length) {
        for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
            var selectedItem = selection[selectionIndex];
            if (!selectedItem || selectedItem.typename !== "TextFrame") continue;

            targets.push({
                textRange: selectedItem.textRange,
                textFrame: selectedItem
            });
        }
    }

    return targets;
}

/* テキスト範囲から親TextFrameを取得 / Get parent TextFrame from text range */
// テキスト編集モードでは textRange.parent が Story を返すことがあるため
// story.textFrames 経由で TextFrame を取得する
function getParentTextFrameFromTextRange(textRange) {
    if (!textRange) return null;
    try {
        var story = textRange.story;
        if (story && story.textFrames && story.textFrames.length > 0) {
            return story.textFrames[0];
        }
    } catch (e) { }
    var parentItem = textRange.parent;
    while (parentItem && parentItem.typename !== "TextFrame") {
        parentItem = parentItem.parent;
    }
    return parentItem || null;
}

/* 段落ごとの行揃えを退避 / Capture justification for each paragraph */
function captureParagraphJustifications(textRange) {
    var justifications = [];
    if (!textRange) return justifications;

    try {
        if (textRange.paragraphs && textRange.paragraphs.length > 0) {
            for (var paragraphIndex = 0; paragraphIndex < textRange.paragraphs.length; paragraphIndex++) {
                justifications.push(textRange.paragraphs[paragraphIndex].paragraphAttributes.justification);
            }
        }
    } catch (e) {
        try {
            justifications.push(textRange.paragraphAttributes.justification);
        } catch (e2) {
        }
    }

    return justifications;
}

/* 段落ごとの行揃えを復元 / Restore justification for each paragraph */
function restoreParagraphJustifications(textRange, textFrame, justifications) {
    if (!textRange || !justifications || justifications.length === 0) return;

    try {
        if (textRange.paragraphs && textRange.paragraphs.length > 0) {
            for (var paragraphIndex = 0; paragraphIndex < textRange.paragraphs.length; paragraphIndex++) {
                var sourceIndex = (paragraphIndex < justifications.length) ? paragraphIndex : justifications.length - 1;
                var originalJustification = justifications[sourceIndex];

                try {
                    textRange.paragraphs[paragraphIndex].paragraphAttributes.justification = originalJustification;
                } catch (paragraphAttrError) {
                }

                try {
                    textRange.paragraphs[paragraphIndex].justification = originalJustification;
                } catch (paragraphDirectError) {
                }
            }
        }
    } catch (e) {
        // 段落単位で復元できない場合でも代表値の復元は続行 / Continue restoring the representative value even if paragraph-level restore fails
    }

    /* 最後に代表値をTextRange/TextFrameへも戻す / Finally restore the representative value to TextRange/TextFrame as well */
    applyJustificationSafely(textRange, textFrame, justifications[0]);
}

/* 現在の属性を退避 / Capture current attributes */
// 混在書式時に値が不定になるのを避けるため先頭文字を基準にする /
// Use the first character to avoid ambiguous values when formatting is mixed
function captureCurrentTextAttributes(textRange, textFrame) {
    var firstCharacter = textRange.characters[0];
    var characterAttributes = firstCharacter.characterAttributes;
    var paragraphAttributes = firstCharacter.paragraphAttributes;
    return {
        fontName: characterAttributes.textFont.name,
        size: characterAttributes.size,
        leading: characterAttributes.leading,
        autoLeading: characterAttributes.autoLeading,
        tracking: characterAttributes.tracking,
        kerningMethod: characterAttributes.kerningMethod,
        proportionalMetrics: characterAttributes.proportionalMetrics,
        orientation: textFrame ? textFrame.orientation : null,
        justification: paragraphAttributes.justification,
        paragraphJustifications: captureParagraphJustifications(textRange)
    };
}

/* 退避した属性を復元 / Restore captured attributes */
function restoreTextAttributes(textRange, textFrame, capturedAttributes) {
    if (!capturedAttributes) return;

    var characterAttributes = textRange.characterAttributes;

    try {
        characterAttributes.textFont = app.textFonts.getByName(capturedAttributes.fontName);
    } catch (e) {
        // 復元時にフォントが見つからない場合は無視 / Ignore if the font cannot be found while restoring
    }

    characterAttributes.size = capturedAttributes.size;
    characterAttributes.autoLeading = capturedAttributes.autoLeading;
    if (!capturedAttributes.autoLeading) {
        characterAttributes.leading = capturedAttributes.leading;
    }
    characterAttributes.tracking = capturedAttributes.tracking;




    applyKerningMethodSafely(characterAttributes, capturedAttributes, textRange);
    characterAttributes.proportionalMetrics = capturedAttributes.proportionalMetrics;
    if (capturedAttributes.paragraphJustifications && capturedAttributes.paragraphJustifications.length > 0) {
        restoreParagraphJustifications(textRange, textFrame, capturedAttributes.paragraphJustifications);
    }
    if (typeof capturedAttributes.justification !== "undefined" && capturedAttributes.justification !== null) {
        applyJustificationSafely(textRange, textFrame, capturedAttributes.justification);
    }
    if (textFrame && capturedAttributes.orientation !== null && typeof capturedAttributes.orientation !== "undefined") {
        textFrame.orientation = capturedAttributes.orientation;
    }
}

/* 複数対象の現在属性を退避 / Capture current attributes for multiple targets */
function captureCurrentTextAttributesForTargets(targets) {
    var capturedList = [];
    for (var targetIndex = 0; targetIndex < targets.length; targetIndex++) {
        capturedList.push(captureCurrentTextAttributes(targets[targetIndex].textRange, targets[targetIndex].textFrame));
    }
    return capturedList;
}

/* 複数対象の退避属性を復元 / Restore captured attributes for multiple targets */
function restoreTextAttributesForTargets(targets, capturedList) {
    if (!targets || !capturedList) return;
    for (var targetIndex = 0; targetIndex < targets.length; targetIndex++) {
        restoreTextAttributes(targets[targetIndex].textRange, targets[targetIndex].textFrame, capturedList[targetIndex]);
    }
}

/* 自動カーニングがメトリクスか判定 / Check whether auto kerning is Metrics */
function isMetricsKerningMethod(copiedAttributes) {
    if (!copiedAttributes) return false;

    try {
        if (copiedAttributes.kerningMethod === AutoKernType.AUTO) return true;
    } catch (e) {
    }

    var kerningMethodLabel = copiedAttributes.kerningMethodLabel;
    if (typeof kerningMethodLabel === "string") {
        return kerningMethodLabel === "メトリクス" || kerningMethodLabel === "Metrics";
    }

    return false;
}

/* 自動カーニングがオプティカルか判定 / Check whether auto kerning is Optical */
function isOpticalKerningMethod(copiedAttributes) {
    if (!copiedAttributes) return false;

    try {
        if (copiedAttributes.kerningMethod === AutoKernType.OPTICAL) return true;
    } catch (e) {
    }

    var kerningMethodLabel = copiedAttributes.kerningMethodLabel;
    if (typeof kerningMethodLabel === "string") {
        return kerningMethodLabel === "オプティカル" || kerningMethodLabel === "Optical";
    }

    return false;
}

/* 自動カーニングが0か判定 / Check whether auto kerning is 0 */
function isZeroKerningMethod(copiedAttributes) {
    if (!copiedAttributes) return false;

    var kerningMethodLabel = copiedAttributes.kerningMethodLabel;
    if (typeof kerningMethodLabel === "string") {
        return kerningMethodLabel === "なし" || kerningMethodLabel === "None";
    }

    try {
        return copiedAttributes.kerningMethod === AutoKernType.NOAUTOKERN;
    } catch (e) {
    }

    return false;
}

/* 現在の選択を配列として退避 / Capture current selection as an array */
function captureSelectionForRestore() {
    var selectionItems = [];

    try {
        if (app.selection && app.selection.length) {
            for (var selectionIndex = 0; selectionIndex < app.selection.length; selectionIndex++) {
                selectionItems.push(app.selection[selectionIndex]);
            }
        } else if (app.selection) {
            selectionItems.push(app.selection);
        }
    } catch (e) {
    }

    return selectionItems;
}

/* 退避した選択を復元 / Restore captured selection */
function restoreSelectionFromCapture(selectionItems) {
    try {
        app.selection = null;
    } catch (e) {
    }

    if (!selectionItems || selectionItems.length === 0) return;

    try {
        app.selection = selectionItems;
    } catch (e2) {
        try {
            app.selection = selectionItems[0];
        } catch (e3) {
        }
    }
}

/* アクション適用前に対象テキスト範囲を選択 / Select target text range before applying an action */
function selectTextRangeForAction(textRange) {
    if (!textRange) return false;

    try {
        textRange.select();
        return true;
    } catch (e) {
    }

    try {
        app.selection = textRange;
        return true;
    } catch (e2) {
    }

    return false;
}

/* 選択を退避・復元しながらアクションを実行 / Run an action while preserving the original selection */
function runKerningActionWithSelection(textRange, actionFunction) {
    if (!actionFunction) return;

    var originalSelection = captureSelectionForRestore();
    try {
        selectTextRangeForAction(textRange);
        actionFunction();
    } finally {
        restoreSelectionFromCapture(originalSelection);
    }
}

/* 自動カーニングを安全に適用 / Apply auto kerning safely */
function applyKerningMethodSafely(characterAttributes, copiedAttributes, textRange) {
    if (!characterAttributes || !copiedAttributes) return;

    if ((typeof copiedAttributes.kerningMethod === "undefined" || copiedAttributes.kerningMethod === null) &&
        (typeof copiedAttributes.kerningMethodLabel !== "string" || copiedAttributes.kerningMethodLabel.length === 0)) return;

    /* Illustrator ExtendScriptでは自動カーニング／メトリクス・オプティカル・0を取得できても直接適用できないため、対象範囲を選択してからアクションで適用する必要がある / In Illustrator ExtendScript, auto kerning Metrics, Optical, and 0 cannot be applied directly even if they can be read, so it is necessary to select the target range and apply it with an action. */
    if (isMetricsKerningMethod(copiedAttributes)) {
        try {
            runKerningActionWithSelection(textRange, runKerningMetricsAction);
        } catch (actionError) {
            // アクションでメトリクスを適用できない場合はスキップ / Skip if Metrics cannot be applied by action
        }
        return;
    }

    if (isOpticalKerningMethod(copiedAttributes)) {
        try {
            runKerningActionWithSelection(textRange, runKerningOpticalAction);
        } catch (actionError2) {
            // アクションでオプティカルを適用できない場合はスキップ / Skip if Optical cannot be applied by action
        }
        return;
    }

    if (isZeroKerningMethod(copiedAttributes)) {
        try {
            runKerningActionWithSelection(textRange, runKerningZeroAction);
        } catch (actionError3) {
            // アクションで0を適用できない場合はスキップ / Skip if 0 cannot be applied by action
        }
        return;
    }

    try {
        characterAttributes.kerningMethod = copiedAttributes.kerningMethod;
    } catch (e) {
        // カーニング方式が現在の文字範囲に適用できない場合はスキップ / Skip if the kerning method cannot be applied to the current text range
    }
}

/* 自動カーニング／メトリクス用アクションを実行 / Run auto kerning Metrics action */
function runKerningMetricsAction() {
    var str = '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 7' + ' 4d657472696373' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1635019621' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 0' + ' }' + ' }' + '}';

    var actionFile = new File('~/ScriptAction.aia');
    try {
        actionFile.open('w');
        actionFile.write(str);
        actionFile.close();
        app.loadAction(actionFile);
        app.doScript('Metrics', 'Kerning', false);
    } finally {
        try {
            app.unloadAction('Kerning', '');
        } catch (unloadError) {
        }
        try {
            if (actionFile.exists) actionFile.remove();
        } catch (removeError) {
        }
    }
}

/* 自動カーニング／オプティカル用アクションを実行 / Run auto kerning Optical action */
function runKerningOpticalAction() {
    var str = '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 7' + ' 4f70746963616c' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1869638501' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 1' + ' }' + ' }' + '}';

    var actionFile = new File('~/ScriptAction.aia');
    try {
        actionFile.open('w');
        actionFile.write(str);
        actionFile.close();
        app.loadAction(actionFile);
        app.doScript('Optical', 'Kerning', false);
    } finally {
        try {
            app.unloadAction('Kerning', '');
        } catch (unloadError) {
        }
        try {
            if (actionFile.exists) actionFile.remove();
        } catch (removeError) {
        }
    }
}

/* 自動カーニング／0用アクションを実行 / Run auto kerning 0 action */
function runKerningZeroAction() {
    var str = '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 8' + ' 4b65726e696e6730' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1801810542' + ' /showInPalette 4294967295' + ' /type (integer)' + ' /value 0' + ' }' + ' }' + '}';

    var actionFile = new File('~/ScriptAction.aia');
    try {
        actionFile.open('w');
        actionFile.write(str);
        actionFile.close();
        app.loadAction(actionFile);
        app.doScript('Kerning0', 'Kerning', false);
    } finally {
        try {
            app.unloadAction('Kerning', '');
        } catch (unloadError) {
        }
        try {
            if (actionFile.exists) actionFile.remove();
        } catch (removeError) {
        }
    }
}


/* 行揃えを安全に適用 / Apply justification safely */
function applyJustificationSafely(textRange, textFrame, justification) {
    if (!textRange) return false;
    if (typeof justification === "undefined" || justification === null) return false;

    try {
        textRange.paragraphAttributes.justification = justification;
        return true;
    } catch (e) {
        // textRangeへ直接適用できない場合は段落単位へフォールバック / Fall back to paragraph-level apply
    }

    var applied = false;
    try {
        if (textRange.paragraphs && textRange.paragraphs.length > 0) {
            for (var paragraphIndex = 0; paragraphIndex < textRange.paragraphs.length; paragraphIndex++) {
                try {
                    textRange.paragraphs[paragraphIndex].paragraphAttributes.justification = justification;
                    applied = true;
                } catch (paragraphError) {
                    try {
                        textRange.paragraphs[paragraphIndex].justification = justification;
                        applied = true;
                    } catch (paragraphDirectError) {
                    }
                }
            }
        }
    } catch (e2) {
        // 段落単位で適用できない場合はTextFrame全体へフォールバック / Fall back to the whole TextFrame
    }

    if (applied) return true;

    try {
        if (textFrame && textFrame.textRange) {
            textFrame.textRange.paragraphAttributes.justification = justification;
            return true;
        }
    } catch (e3) {
        // 行揃えが現在の範囲に適用できない場合はスキップ / Skip if justification cannot be applied to the current range
    }

    return false;
}

/* コピー済み属性を適用 / Apply copied attributes */
function applyCopiedTextAttributes(textRange, textFrame, copiedAttributes, uiState) {
    var characterAttributes = textRange.characterAttributes;

    if (uiState.applyFont) {
        try {
            var targetFont = app.textFonts.getByName(copiedAttributes.font.name);
            characterAttributes.textFont = targetFont;
        } catch (e) {
            throw new Error(L("errorFontNotFoundPrefix") + copiedAttributes.font.name + L("errorFontNotFoundSuffix"));
        }
    }

    if (uiState.applySize) {
        characterAttributes.size = copiedAttributes.size;
    }

    if (uiState.applyAutoLeading) {
        characterAttributes.autoLeading = copiedAttributes.autoLeading;
    }

    if (uiState.applyLeading && !copiedAttributes.autoLeading) {
        characterAttributes.leading = copiedAttributes.leading;
    }

    if (uiState.applyTracking) {
        characterAttributes.tracking = copiedAttributes.tracking;
    }

    if (uiState.applyKerningMethod) {
        applyKerningMethodSafely(characterAttributes, copiedAttributes, textRange);
    }

    if (uiState.applyProportionalMetrics) {
        characterAttributes.proportionalMetrics = copiedAttributes.proportionalMetrics;
    }

    if (uiState.applyJustification && typeof copiedAttributes.justification !== "undefined" && copiedAttributes.justification !== null) {
        applyJustificationSafely(textRange, textFrame, copiedAttributes.justification);
    }

    if (uiState.applyOrientation && textFrame && copiedAttributes.orientation !== null) {
        textFrame.orientation = copiedAttributes.orientation;
    }
}

/* 複数対象へコピー済み属性を適用 / Apply copied attributes to multiple targets */
function applyCopiedTextAttributesToTargets(targets, copiedAttributes, uiState) {
    for (var targetIndex = 0; targetIndex < targets.length; targetIndex++) {
        applyCopiedTextAttributes(targets[targetIndex].textRange, targets[targetIndex].textFrame, copiedAttributes, uiState);
    }
}

/* UI状態を取得 / Read UI state */
function readApplyUIState(ui) {
    return {
        applyFont: ui.cbFont.value,
        applySize: ui.cbSize.value,
        applyLeading: ui.cbLeading.value,
        applyAutoLeading: ui.cbAutoLeading.value,
        applyTracking: ui.cbTracking.value,
        applyKerningMethod: ui.cbKerningMethod.value,
        applyProportionalMetrics: ui.cbProportionalMetrics.value,
        applyOrientation: ui.cbOrientation.value,
        applyJustification: ui.cbJustification.value
    };
}

/* 適用対象があるか確認 / Check whether any item should be applied */
function hasAnyApplyTarget(uiState) {
    return uiState.applyFont ||
        uiState.applySize ||
        uiState.applyLeading ||
        uiState.applyAutoLeading ||
        uiState.applyTracking ||
        uiState.applyKerningMethod ||
        uiState.applyProportionalMetrics ||
        uiState.applyOrientation ||
        uiState.applyJustification;
}

/* コピー済み属性の有無を確認 / Check whether copied attributes exist */
function hasCopiedFont(copiedAttributes) {
    return !!(
        copiedAttributes &&
        copiedAttributes.font &&
        typeof copiedAttributes.font.name === "string" &&
        copiedAttributes.font.name.length > 0
    );
}

function hasCopiedSize(copiedAttributes) {
    return !!(
        copiedAttributes &&
        typeof copiedAttributes.size === "number" &&
        !isNaN(copiedAttributes.size) &&
        copiedAttributes.size > 0
    );
}

function hasCopiedNumber(copiedAttributes, key) {
    return !!(
        copiedAttributes &&
        typeof copiedAttributes[key] === "number" &&
        !isNaN(copiedAttributes[key])
    );
}

function hasCopiedBoolean(copiedAttributes, key) {
    return !!(
        copiedAttributes &&
        typeof copiedAttributes[key] === "boolean"
    );
}

function hasCopiedOrientation(copiedAttributes) {
    return !!(
        copiedAttributes &&
        copiedAttributes.orientation !== null &&
        typeof copiedAttributes.orientation !== "undefined"
    );
}

function hasCopiedJustification(copiedAttributes) {
    return !!(
        copiedAttributes &&
        copiedAttributes.justification !== null &&
        typeof copiedAttributes.justification !== "undefined"
    );
}

/* コピー済みの自動カーニングがあるか確認 / Check whether copied auto kerning exists */
function hasCopiedKerningMethod(copiedAttributes) {
    if (!copiedAttributes) return false;

    if (copiedAttributes.kerningMethod !== null && typeof copiedAttributes.kerningMethod !== "undefined") {
        return true;
    }

    return !!(
        typeof copiedAttributes.kerningMethodLabel === "string" &&
        copiedAttributes.kerningMethodLabel.length > 0
    );
}

/* ON/OFF表示を返す / Return ON/OFF display text */
function formatBooleanForDialog(value) {
    return value ? L("onValue") : L("offValue");
}

/* 組み方向表示を返す / Return orientation display text */
function formatOrientationForDialog(orientation) {
    if (orientation === TextOrientation.VERTICAL) {
        return L("verticalOrientation");
    }
    return L("horizontalOrientation");
}

/* ラベル幅を固定したチェックボックス行を追加 / Add checkbox row with fixed label width */
function addAttributeCheckboxRow(parent, labelKey, valueText, isAvailable, defaultValue, labelWidth) {
    var rowGroup = parent.add("group");
    rowGroup.orientation = "row";
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = 6;

    var label = rowGroup.add("statictext", undefined, labelText(labelKey));
    label.preferredSize.width = labelWidth;
    label.justify = "right";

    var checkboxLabel = isAvailable ? valueText : L("notStored");
    var checkbox = rowGroup.add("checkbox", undefined, checkboxLabel);
    checkbox.value = isAvailable && defaultValue;
    checkbox.enabled = isAvailable;
    return checkbox;
}

/* フォントとスタイルを2行で表示するチェックボックス行を追加 / Add checkbox row that displays font and style in two lines */
function addFontCheckboxRow(parent, fontText, styleText, isAvailable, defaultValue, labelWidth) {
    var CHECKBOX_TEXT_OFFSET = 22;

    var rowGroup = parent.add("group");
    rowGroup.orientation = "row";
    rowGroup.alignChildren = ["left", "top"];
    rowGroup.spacing = 6;

    var label = rowGroup.add("statictext", undefined, labelText("fontLabel"));
    label.preferredSize.width = labelWidth;
    label.justify = "right";

    var valueGroup = rowGroup.add("group");
    valueGroup.orientation = "column";
    valueGroup.alignChildren = "left";
    valueGroup.spacing = 2;

    var checkboxLabel = isAvailable ? fontText : L("notStored");
    var checkbox = valueGroup.add("checkbox", undefined, checkboxLabel);
    checkbox.value = isAvailable && defaultValue;
    checkbox.enabled = isAvailable;

    if (isAvailable) {
        var styleRow = valueGroup.add("group");
        styleRow.orientation = "row";
        styleRow.alignChildren = ["left", "center"];
        styleRow.spacing = 0;
        styleRow.margins = [0, 5, 0, 0];

        var spacer = styleRow.add("statictext", undefined, "");
        spacer.preferredSize.width = CHECKBOX_TEXT_OFFSET;

        styleRow.add("statictext", undefined, styleText);
    }

    return checkbox;
}


/* Optionクリックで対象以外をOFFにする / Turn off other checkboxes with Option-click */
function bindExclusiveOptionClick(checkboxes) {
    for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
        (function (targetCheckbox) {
            targetCheckbox.onClick = function () {
                if (!ScriptUI.environment.keyboardState.altKey) return;

                for (var i = 0; i < checkboxes.length; i++) {
                    if (!checkboxes[i].enabled) continue;
                    checkboxes[i].value = (checkboxes[i] === targetCheckbox);
                }
            };
        })(checkboxes[checkboxIndex]);
    }
}

// =========================================
// メイン処理 / Main process
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(L("errorNoDocument"));
        return;
    }

    /* 同じ #targetengine の $.global.FontClipboard から取得 / Read from $.global.FontClipboard in the same #targetengine */
    var copiedTextAttributes = $.global.FontClipboard;

    var hasFont = hasCopiedFont(copiedTextAttributes);
    var hasSize = hasCopiedSize(copiedTextAttributes);
    var hasLeading = hasCopiedNumber(copiedTextAttributes, "leading");
    var hasAutoLeading = hasCopiedBoolean(copiedTextAttributes, "autoLeading");
    var hasTracking = hasCopiedNumber(copiedTextAttributes, "tracking");
    var hasKerningMethod = hasCopiedKerningMethod(copiedTextAttributes);
    var hasProportionalMetrics = hasCopiedBoolean(copiedTextAttributes, "proportionalMetrics");
    var hasOrientation = hasCopiedOrientation(copiedTextAttributes);
    var hasJustification = hasCopiedJustification(copiedTextAttributes);

    if (!hasFont && !hasSize && !hasLeading && !hasAutoLeading && !hasTracking && !hasKerningMethod && !hasProportionalMetrics && !hasOrientation && !hasJustification) {
        alert(L("errorNoCopiedAttributes"));
        return;
    }

    var applyTargets = getApplyTargetsFromSelection(app.selection);

    if (applyTargets.length === 0) {
        alert(L("errorNoTextSelection"));
        return;
    }

    for (var applyTargetIndex = 0; applyTargetIndex < applyTargets.length; applyTargetIndex++) {
        if (applyTargets[applyTargetIndex].textRange.characters.length === 0) {
            alert(L("errorEmptyTextRange"));
            return;
        }
    }

    var canApplyOrientation = true;
    for (var orientationTargetIndex = 0; orientationTargetIndex < applyTargets.length; orientationTargetIndex++) {
        if (!applyTargets[orientationTargetIndex].textFrame) {
            canApplyOrientation = false;
            break;
        }
    }

    /* ダイアログ / Dialog */
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.alignChildren = "left";
    dlg.margins = [16, 16, 26, 16];
    dlg.spacing = 10;

    var fontDisplay = hasFont ? (copiedTextAttributes.font.family || L("notStored")) : L("notStored");
    var fontStyleDisplay = hasFont ? (copiedTextAttributes.font.style || L("notStored")) : L("notStored");
    var sizeDisplay = hasSize ? formatPointValueForDialog(copiedTextAttributes.size) : L("notStored");
    var leadingDisplay = hasLeading ? formatPointValueForDialog(copiedTextAttributes.leading) : L("notStored");
    var autoLeadingDisplay = hasAutoLeading ? formatBooleanForDialog(copiedTextAttributes.autoLeading) : L("notStored");
    var trackingDisplay = hasTracking ? String(copiedTextAttributes.tracking) : L("notStored");
    var kerningMethodDisplay = hasKerningMethod
        ? (copiedTextAttributes.kerningMethodLabel || String(copiedTextAttributes.kerningMethod))
        : L("notStored");
    var proportionalMetricsDisplay = hasProportionalMetrics ? formatBooleanForDialog(copiedTextAttributes.proportionalMetrics) : L("notStored");
    var orientationDisplay = hasOrientation
        ? (copiedTextAttributes.orientationLabel || formatOrientationForDialog(copiedTextAttributes.orientation))
        : L("notStored");
    var justificationDisplay = hasJustification
        ? (copiedTextAttributes.justificationLabel || String(copiedTextAttributes.justification))
        : L("notStored");

    var cbFont = addFontCheckboxRow(dlg, fontDisplay, fontStyleDisplay, hasFont, true, ATTRIBUTE_LABEL_WIDTH);
    var cbSize = addAttributeCheckboxRow(dlg, "sizeLabel", sizeDisplay, hasSize, false, ATTRIBUTE_LABEL_WIDTH);
    var cbLeading = addAttributeCheckboxRow(dlg, "leadingLabel", leadingDisplay, hasLeading, false, ATTRIBUTE_LABEL_WIDTH);
    var cbAutoLeading = addAttributeCheckboxRow(dlg, "autoLeadingLabel", autoLeadingDisplay, hasAutoLeading, false, ATTRIBUTE_LABEL_WIDTH);
    var cbTracking = addAttributeCheckboxRow(dlg, "trackingLabel", trackingDisplay, hasTracking, false, ATTRIBUTE_LABEL_WIDTH);
    var cbKerningMethod = addAttributeCheckboxRow(dlg, "kerningMethodLabel", kerningMethodDisplay, hasKerningMethod, false, ATTRIBUTE_LABEL_WIDTH);
    var cbProportionalMetrics = addAttributeCheckboxRow(dlg, "proportionalMetricsLabel", proportionalMetricsDisplay, hasProportionalMetrics, false, ATTRIBUTE_LABEL_WIDTH);
    var cbOrientation = addAttributeCheckboxRow(dlg, "orientationLabel", orientationDisplay, hasOrientation, false, ATTRIBUTE_LABEL_WIDTH);
    cbOrientation.enabled = hasOrientation && canApplyOrientation;
    var cbJustification = addAttributeCheckboxRow(dlg, "justificationLabel", justificationDisplay, hasJustification, false, ATTRIBUTE_LABEL_WIDTH);

    var attributeCheckboxes = [
        cbFont,
        cbSize,
        cbLeading,
        cbAutoLeading,
        cbTracking,
        cbKerningMethod,
        cbProportionalMetrics,
        cbOrientation,
        cbJustification
    ];
    bindExclusiveOptionClick(attributeCheckboxes);

    /* ボタンエリア（cbPreviewを含む）を先に生成し、updatePreviewから安全に参照できるようにする
    
       Build the button area (including cbPreview) first so updatePreview can reference it safely */

    var buttonArea = dlg.add("group");
    buttonArea.orientation = "row";
    buttonArea.alignChildren = ["fill", "center"];
    buttonArea.alignment = "fill";
    buttonArea.margins = [0, 10, 0, 0];

    var previewGroup = buttonArea.add("group");
    previewGroup.orientation = "row";
    previewGroup.alignChildren = ["left", "center"];
    var cbPreview = previewGroup.add("checkbox", undefined, L("previewCheckbox"));
    cbPreview.value = false;

    var spacer = buttonArea.add("group");
    spacer.alignment = ["fill", "center"];
    spacer.minimumSize.width = 20;

    var btnGroup = buttonArea.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = ["right", "center"];
    btnGroup.add("button", undefined, L("cancelButton"), { name: "cancel" });
    btnGroup.add("button", undefined, L("applyButton"), { name: "ok" });

    var originalTextAttributesList = captureCurrentTextAttributesForTargets(applyTargets);
    var isPreviewApplied = false;

    function removePreview() {
        if (!isPreviewApplied) return;
        restoreTextAttributesForTargets(applyTargets, originalTextAttributesList);
        isPreviewApplied = false;
    }

    function updatePreview() {
        if (!cbPreview.value) {
            removePreview();
            return;
        }

        try {
            removePreview();
            var previewState = readApplyUIState({
                cbFont: cbFont,
                cbSize: cbSize,
                cbLeading: cbLeading,
                cbAutoLeading: cbAutoLeading,
                cbTracking: cbTracking,
                cbKerningMethod: cbKerningMethod,
                cbProportionalMetrics: cbProportionalMetrics,
                cbOrientation: cbOrientation,
                cbJustification: cbJustification
            });
            applyCopiedTextAttributesToTargets(applyTargets, copiedTextAttributes, previewState);
            isPreviewApplied = true;
            app.redraw();
        } catch (e) {
            restoreTextAttributesForTargets(applyTargets, originalTextAttributesList);
            isPreviewApplied = false;
            cbPreview.value = false;
            alert(e.message);
        }
    }

    cbPreview.onClick = updatePreview;

    for (var previewCheckboxIndex = 0; previewCheckboxIndex < attributeCheckboxes.length; previewCheckboxIndex++) {
        attributeCheckboxes[previewCheckboxIndex].onClick = (function (originalOnClick) {
            return function () {
                if (originalOnClick) originalOnClick.call(this);
                updatePreview();
            };
        })(attributeCheckboxes[previewCheckboxIndex].onClick);
    }

    var dialogResult = dlg.show();
    if (dialogResult !== 1) {
        restoreTextAttributesForTargets(applyTargets, originalTextAttributesList);
        isPreviewApplied = false;
        return;
    }

    var applyState = readApplyUIState({
        cbFont: cbFont,
        cbSize: cbSize,
        cbLeading: cbLeading,
        cbAutoLeading: cbAutoLeading,
        cbTracking: cbTracking,
        cbKerningMethod: cbKerningMethod,
        cbProportionalMetrics: cbProportionalMetrics,
        cbOrientation: cbOrientation,
        cbJustification: cbJustification
    });

    if (!hasAnyApplyTarget(applyState)) {
        restoreTextAttributesForTargets(applyTargets, originalTextAttributesList);
        isPreviewApplied = false;
        alert(L("errorNoApplyItem"));
        return;
    }

    /* 適用 / Apply */
    // textRange 全体に対して characterAttributes を書き換える / Rewrite characterAttributes for the entire textRange
    // プレビューONで表示済みの状態は適用内容と一致するため、戻す→再適用はスキップ /
    // When preview is ON, the on-screen state already matches applyState, so skip the revert/re-apply cycle
    try {
        if (isPreviewApplied && cbPreview.value) {
            isPreviewApplied = false;
        } else {
            removePreview();
            applyCopiedTextAttributesToTargets(applyTargets, copiedTextAttributes, applyState);
        }
    } catch (e) {
        restoreTextAttributesForTargets(applyTargets, originalTextAttributesList);
        isPreviewApplied = false;
        alert(e.message);
        return;
    }
})();