#target illustrator
#targetengine "FontClipboard"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択されたテキストオブジェクト、またはテキスト編集モードで部分選択された文字列から、
先頭文字を基準にフォント情報・文字属性・段落属性を取得し、永続エンジン "FontClipboard" の
$.global.FontClipboard に保存します。
保存する内容は、フォントの PostScript 名・フォントファミリ名・スタイル・フォントサイズ・行送り・自動行送り・
トラッキング・自動カーニング・自動カーニングの表示名・プロポーショナルメトリクス・
組み方向・組み方向の表示名・行揃え・行揃えの表示名です。
フォントサイズは Illustrator 内部値の pt で保存し、完了メッセージでは text/units の設定に合わせて表示します。
同じ #targetengine を指定した適用用スクリプトから値を読み取れます。

Overview

This script reads font information, character attributes, and paragraph attributes from the first character
of the selected text object or partially selected text range, then stores them in $.global.FontClipboard
on the persistent "FontClipboard" engine.
The stored values include the font PostScript name, font family name, style, font size, leading,
auto leading, tracking, auto kerning, auto kerning label, proportional metrics,
text orientation, orientation label, alignment, and alignment label.
Font size is stored as Illustrator's internal point value, while the completion message displays it
according to the text/units preference.
A separate apply script using the same #targetengine can read the saved values.

作成日 / Created: 2026-04-25
*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLocaleLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLocaleLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    errorNoDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    errorNoTextSelection: {
        ja: "テキストオブジェクトを 1 つ選択するか、テキスト編集モードで文字列を選択してください。",
        en: "Select one text object, or select text in text editing mode."
    },
    errorEmptyTextRange: {
        ja: "文字のある範囲を選択してください。",
        en: "Select a text range that contains characters."
    },
    copiedMessageTitle: {
        ja: "フォント情報を取得しました",
        en: "Font information copied"
    },
    postScriptName: {
        ja: "PostScript名",
        en: "PostScript name"
    },
    fontFamilyName: {
        ja: "フォントファミリ名",
        en: "Font family name"
    },
    fontStyle: {
        ja: "スタイル",
        en: "Style"
    },
    fontSize: {
        ja: "フォントサイズ",
        en: "Font size"
    },
    leading: {
        ja: "行送り",
        en: "Leading"
    },
    autoLeading: {
        ja: "自動行送り",
        en: "Auto leading"
    },
    tracking: {
        ja: "トラッキング",
        en: "Tracking"
    },
    kerningMethod: {
        ja: "カーニング",
        en: "Kerning"
    },
    proportionalMetrics: {
        ja: "プロポーショナルメトリクス",
        en: "Proportional metrics"
    },
    orientation: {
        ja: "組み方向",
        en: "Orientation"
    },
    justification: {
        ja: "行揃え",
        en: "Alignment"
    },
    valueOn: {
        ja: "オン",
        en: "On"
    },
    valueOff: {
        ja: "オフ",
        en: "Off"
    },
    closeButton: {
        ja: "閉じる",
        en: "Close"
    }
};

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
function formatPointValueForDisplay(pointValue) {
    var unitCode = getPreferenceUnitCode("text/units", 2);
    var textUnitLabel = getPreferenceUnitLabel(unitCode, "text/units");
    var displayValue = convertPointsToPreferenceUnit(pointValue, unitCode);
    return formatNumberForDisplay(displayValue, 3) + " " + textUnitLabel;
}

/* Boolean値をオン/オフラベルに変換 / Convert boolean to on/off label */
function formatBooleanLabel(boolValue) {
    return boolValue ? L("valueOn") : L("valueOff");
}

// =========================================
// 文字属性ユーティリティ / Text attribute utilities
// =========================================

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
    } catch (e) {}
    var parentItem = textRange.parent;
    while (parentItem && parentItem.typename !== "TextFrame") {
        parentItem = parentItem.parent;
    }
    return parentItem || null;
}

/* カーニング方式の表示名を取得 / Get display label for kerning method */
function getKerningMethodLabel(kerningMethod) {
    switch (kerningMethod) {
        case AutoKernType.AUTO:
            return (lang === "ja") ? "メトリクス" : "Metrics";
        case AutoKernType.METRICSROMANONLY:
            return (lang === "ja") ? "和文等幅" : "Metrics - Roman Only";
        case AutoKernType.OPTICAL:
            return (lang === "ja") ? "オプティカル" : "Optical";
        default:
            return (lang === "ja") ? "なし" : "None";
    }
}

/* 組み方向の表示名を取得 / Get display label for text orientation */
function getOrientationLabel(textFrame) {
    if (!textFrame) return "";
    return (textFrame.orientation === TextOrientation.VERTICAL)
        ? ((lang === "ja") ? "縦組み" : "Vertical")
        : ((lang === "ja") ? "横組み" : "Horizontal");
}

/* 結果表示ダイアログ / Show result dialog */
function showResultDialog(info) {
    var dialog = new Window("dialog", L("copiedMessageTitle"));
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 16;

    var infoPanel = dialog.add("panel", undefined, "");
    infoPanel.orientation = "column";
    infoPanel.alignChildren = "left";
    infoPanel.margins = 14;
    infoPanel.spacing = 4;

    function addRow(key, value) {
        var row = infoPanel.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = 8;
        var labelControl = row.add("statictext", undefined, labelText(key));
        labelControl.preferredSize.width = 180;
        row.add("statictext", undefined, String(value));
    }

    addRow("postScriptName", info.font.name);
    addRow("fontFamilyName", info.font.family);
    addRow("fontStyle", info.font.style);
    addRow("fontSize", formatPointValueForDisplay(info.size));
    addRow("leading", formatPointValueForDisplay(info.leading));
    addRow("autoLeading", formatBooleanLabel(info.autoLeading));
    addRow("tracking", info.tracking);
    addRow("kerningMethod", info.kerningMethodLabel);
    addRow("proportionalMetrics", formatBooleanLabel(info.proportionalMetrics));
    addRow("orientation", info.orientationLabel);
    addRow("justification", info.justificationLabel);

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var closeButton = buttonGroup.add("button", undefined, L("closeButton"), { name: "ok" });
    closeButton.onClick = function () { dialog.close(); };

    dialog.show();
}

/* 行揃えの表示名を取得 / Get display label for paragraph justification */
function getJustificationLabel(justification) {
    switch (justification) {
        case Justification.LEFT:
            return (lang === "ja") ? "左揃え" : "Left";
        case Justification.CENTER:
            return (lang === "ja") ? "中央揃え" : "Center";
        case Justification.RIGHT:
            return (lang === "ja") ? "右揃え" : "Right";
        case Justification.FULLJUSTIFY:
            return (lang === "ja") ? "均等配置（最終行左揃え）" : "Justify with last line aligned left";
        case Justification.FULLJUSTIFYLASTLINECENTER:
            return (lang === "ja") ? "均等配置（最終行中央揃え）" : "Justify with last line centered";
        case Justification.FULLJUSTIFYLASTLINERIGHT:
            return (lang === "ja") ? "均等配置（最終行右揃え）" : "Justify with last line aligned right";
        case Justification.FULLJUSTIFYALL:
            return (lang === "ja") ? "両端揃え" : "Justify all lines";
        default:
            return String(justification);
    }
}

(function () {
    if (app.documents.length === 0) {
        alert(L("errorNoDocument"));
        return;
    }

    var currentSelection = app.selection;
    var sourceTextRange = null;

    /* テキスト編集モードでの部分選択 / Partial text selection in text editing mode */
    // app.selection は TextRange オブジェクトが直接返る / app.selection directly returns a TextRange object
    if (currentSelection && currentSelection.typename === "TextRange") {
        sourceTextRange = currentSelection;
    }
    /* オブジェクト選択モード / Object selection mode */
    // 配列で 1 つの TextFrame が返る / A single TextFrame is returned in an array
    else if (currentSelection && currentSelection.length === 1 && currentSelection[0] && currentSelection[0].typename === "TextFrame") {
        sourceTextRange = currentSelection[0].textRange;
    }

    if (!sourceTextRange) {
        alert(L("errorNoTextSelection"));
        return;
    }

    if (sourceTextRange.characters.length === 0) {
        alert(L("errorEmptyTextRange"));
        return;
    }

    /* 先頭文字の属性を参照 / Read attributes from the first character */
    // 混在書式時に値が不定になるのを避ける / Avoid ambiguous values when formatting is mixed
    var sourceFirstCharacter = sourceTextRange.characters[0];
    var sourceCharacterAttributes = sourceFirstCharacter.characterAttributes;
    var sourceParagraphAttributes = sourceFirstCharacter.paragraphAttributes;

    var copiedLeading = sourceCharacterAttributes.leading;
    var copiedAutoLeading = sourceCharacterAttributes.autoLeading;
    var copiedKerningMethod = sourceCharacterAttributes.kerningMethod;
    var copiedProportionalMetrics = sourceCharacterAttributes.proportionalMetrics;
    var copiedTracking = sourceCharacterAttributes.tracking;
    var copiedJustification = sourceParagraphAttributes.justification;
    var copiedJustificationLabel = getJustificationLabel(copiedJustification);

    var sourceTextFrame = getParentTextFrameFromTextRange(sourceTextRange);
    var copiedOrientation = sourceTextFrame ? sourceTextFrame.orientation : null;
    var copiedOrientationLabel = getOrientationLabel(sourceTextFrame);
    var copiedKerningMethodLabel = getKerningMethodLabel(copiedKerningMethod);

    $.global.FontClipboard = {
        font: {
            name: sourceCharacterAttributes.textFont.name,
            family: sourceCharacterAttributes.textFont.family,
            style: sourceCharacterAttributes.textFont.style
        },
        size: sourceCharacterAttributes.size,
        leading: copiedLeading,
        autoLeading: copiedAutoLeading,
        tracking: copiedTracking,
        kerningMethod: copiedKerningMethod,
        kerningMethodLabel: copiedKerningMethodLabel,
        proportionalMetrics: copiedProportionalMetrics,
        orientation: copiedOrientation,
        orientationLabel: copiedOrientationLabel,
        justification: copiedJustification,
        justificationLabel: copiedJustificationLabel
        // 将来拡張
    };

    showResultDialog($.global.FontClipboard);
})();