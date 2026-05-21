#target illustrator
#targetengine "FontClipboard"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択されたテキストオブジェクト、またはテキスト編集モードで部分選択された文字列から、
先頭文字を基準に文字属性・段落属性を取得し、永続エンジン "FontClipboard" の
$.global.FontClipboard に保存します。同じ #targetengine を指定した
ApplyTextAttributesFromClipboard.jsx から読み取って適用できます。

保存する内容
- フォント（PostScript 名・ファミリ名・スタイル）／フォントサイズ／行送り／自動行送り
- 文字ツメ／トラッキング／自動カーニング（種別と表示名）／プロポーショナルメトリクス
- 組み方向（値と表示名）／行揃え（値と表示名）
- 塗り色（複製）と塗り色の表示名
- グラフィックスタイル名（後述）

グラフィックスタイルの自動登録
- 選択中の TextFrame の見た目を、register-temp-style.jsx と同じダイナミックアクションで
  固定名 "temp_style" としてドキュメントに登録し、その名前を保存します。
- 同名スタイルが存在する場合は削除してから上書きするので、繰り返し実行しても重複しません。
- 登録中だけ TextFrame をオブジェクト選択し直し、完了後に元の選択（TextRange 含む）へ戻します。

表示と単位
- フォントサイズと行送りは Illustrator 内部値（pt）で保存し、結果ダイアログでは
  text/units の設定に合わせて表示します。
- 自動行送りがオンのときは行送りの数値を「—」で表示します。
- 行揃えは Justification 値を判定し、取得できない場合は「—」と表示します。

結果ダイアログにはラジオやチェックボックスは置かず、取得した塗り色とグラフィックスタイル名を
そのまま並べて表示します。Apply 側でどちらを使うかを選びます。

Overview

This script reads character and paragraph attributes from the first character of the selected
TextFrame (or partially selected text range) and stores them in $.global.FontClipboard on the
persistent "FontClipboard" engine. ApplyTextAttributesFromClipboard.jsx, running under the same
#targetengine, reads and applies them.

Stored values
- Font (PostScript name, family, style), size, leading, auto leading
- Tsume, tracking, auto kerning (type and label), proportional metrics
- Text orientation (value and label), alignment (value and label)
- Fill color (cloned) and its display label
- Graphic style name (see below)

Auto-registered graphic style
- The current TextFrame's appearance is registered as a graphic style with the fixed name
  "temp_style" by running the same dynamic action used in register-temp-style.jsx.
- Any existing "temp_style" is removed first, so repeated runs do not create duplicates.
- The TextFrame is temporarily selected as an object during registration; the original selection
  (including a TextRange) is restored afterwards.

Display and units
- Font size and leading are stored as internal point values; the result dialog displays them
  according to the text/units preference.
- When auto leading is on, the leading is shown as "—".
- Alignment is resolved from the Justification value; otherwise shown as "—".

The result dialog only displays the captured values (no radios or checkboxes); the choice between
fill and graphic style is made on the Apply side.

作成日 / Created: 2021-04-10
更新日 / Updated: 2026-05-21
*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.3.0";

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
    fontSizeLeadingPanelTitle: {
        ja: "フォント・サイズ、行送り",
        en: "Font, Size & Leading"
    },
    kerningPanelTitle: {
        ja: "カーニング関連",
        en: "Kerning"
    },
    paragraphOtherPanelTitle: {
        ja: "段落属性ほか",
        en: "Paragraph & Other"
    },
    fillGraphicStylePanelTitle: {
        ja: "塗りとグラフィックスタイル",
        en: "Fill & Graphic Style"
    },
    graphicStyle: {
        ja: "グラフィックスタイル",
        en: "Graphic Style"
    },
    graphicStyleNotRegistered: {
        ja: "—",
        en: "—"
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
    tsume: {
        ja: "文字ツメ",
        en: "Tsume"
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
    fillColor: {
        ja: "塗り",
        en: "Fill"
    },
    fillColorNone: {
        ja: "なし",
        en: "None"
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
// 塗り色ユーティリティ / Fill color utilities
// =========================================

/* 色を安全に複製 / Safely clone a color */
// 参照のまま保持すると Illustrator 側の状態に引きずられるため、新規インスタンスを返す
function cloneColor(color) {
    if (!color) return null;
    switch (color.typename) {
        case "RGBColor":
            var rgb = new RGBColor();
            rgb.red = color.red;
            rgb.green = color.green;
            rgb.blue = color.blue;
            return rgb;
        case "CMYKColor":
            var cmyk = new CMYKColor();
            cmyk.cyan = color.cyan;
            cmyk.magenta = color.magenta;
            cmyk.yellow = color.yellow;
            cmyk.black = color.black;
            return cmyk;
        case "GrayColor":
            var gray = new GrayColor();
            gray.gray = color.gray;
            return gray;
        case "SpotColor":
            var spot = new SpotColor();
            spot.spot = color.spot;
            spot.tint = color.tint;
            return spot;
        case "GradientColor":
            var grad = new GradientColor();
            grad.gradient = color.gradient;
            grad.angle = color.angle;
            grad.length = color.length;
            grad.origin = color.origin;
            grad.matrix = color.matrix;
            return grad;
        case "NoColor":
            return new NoColor();
        default:
            return null;
    }
}

/* テキスト範囲から塗り色を取得 / Get fill color from a text range */
function getTextRangeFillColor(textRange) {
    try {
        return textRange.characterAttributes.fillColor;
    } catch (e) {
        return null;
    }
}

/* 色を表示用文字列へ整形 / Format color for display */
function formatColorForDisplay(color) {
    if (!color || !color.typename) return L("fillColorNone");
    switch (color.typename) {
        case "RGBColor":
            return "RGB(" + Math.round(color.red) + ", " + Math.round(color.green) + ", " + Math.round(color.blue) + ")";
        case "CMYKColor":
            return "CMYK(" + Math.round(color.cyan) + ", " + Math.round(color.magenta) + ", " + Math.round(color.yellow) + ", " + Math.round(color.black) + ")";
        case "GrayColor":
            return "Gray(" + Math.round(color.gray) + ")";
        case "SpotColor":
            try {
                return "Spot: " + color.spot.name + " (" + Math.round(color.tint) + "%)";
            } catch (e) {
                return "Spot";
            }
        case "GradientColor":
            try {
                return "Gradient: " + color.gradient.name;
            } catch (e) {
                return "Gradient";
            }
        case "NoColor":
            return L("fillColorNone");
        default:
            return color.typename;
    }
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
    } catch (e) { }
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

/* 自動行送り値を真偽値として解釈 / Interpret auto leading value as boolean */
function isAutoLeadingValueOn(autoLeadingValue) {
    return autoLeadingValue === true || autoLeadingValue === 1;
}

/* 文字範囲の自動行送りを確認 / Check auto leading in a text range */
function hasAutoLeadingInTextRange(textRange) {
    if (!textRange) return false;

    try {
        if (textRange.characterAttributes && isAutoLeadingValueOn(textRange.characterAttributes.autoLeading)) {
            return true;
        }
    } catch (e) {
    }

    try {
        if (textRange.characters && textRange.characters.length > 0) {
            for (var characterIndex = 0; characterIndex < textRange.characters.length; characterIndex++) {
                try {
                    if (isAutoLeadingValueOn(textRange.characters[characterIndex].characterAttributes.autoLeading)) {
                        return true;
                    }
                } catch (characterError) {
                }
            }
        }
    } catch (e2) {
    }

    try {
        if (textRange.lines && textRange.lines.length > 0) {
            for (var lineIndex = 0; lineIndex < textRange.lines.length; lineIndex++) {
                try {
                    if (isAutoLeadingValueOn(textRange.lines[lineIndex].characterAttributes.autoLeading)) {
                        return true;
                    }
                } catch (lineError) {
                }
            }
        }
    } catch (e3) {
    }

    return false;
}

/* 自動行送りかどうかを安全に取得 / Safely detect whether auto leading is enabled */
function getAutoLeadingState(textRange, firstCharacter, textFrame) {
    try {
        if (firstCharacter && firstCharacter.characterAttributes) {
            if (isAutoLeadingValueOn(firstCharacter.characterAttributes.autoLeading)) {
                return true;
            }
        }
    } catch (e) {
    }

    if (hasAutoLeadingInTextRange(textRange)) {
        return true;
    }

    try {
        if (textFrame && textFrame.textRange && hasAutoLeadingInTextRange(textFrame.textRange)) {
            return true;
        }
    } catch (e2) {
    }

    return false;
}

// =========================================
// グラフィックスタイル登録 / Graphic style registration
// =========================================

/* 登録するスタイル名 / Style name to register
   register-temp-style.jsx と同じ固定名。Apply 側はこの名前を取り出して適用する。
   Same fixed name as register-temp-style.jsx; the Apply script reads this name back. */
var TEMP_STYLE_NAME = "temp_style";
var TEMP_STYLE_ACTION_SET = "GraphicStyle";
var TEMP_STYLE_ACTION_NAME = "AddNewWithoutName";

/* 強制的に無名グラフィックスタイルを追加するアクションを書き出してロード /
   Write a dynamic action that appends an unnamed graphic style, and load it */
function loadForceNewGraphicStyleAction() {
    var actionData = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382 a4e383ab ] /value 1 } } }';

    var actionFile = new File(Folder.temp.fsName + '/__tmp_register_style_' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000) + '.aia');
    actionFile.open('w');
    actionFile.write(actionData);
    actionFile.close();
    try {
        app.loadAction(actionFile);
    } finally {
        try {
            if (actionFile.exists) actionFile.remove();
        } catch (removeError) {
        }
    }
}

/* ロード済みアクションを実行 / Run the loaded action */
function runForceNewGraphicStyleAction() {
    app.doScript(TEMP_STYLE_ACTION_NAME, TEMP_STYLE_ACTION_SET, false);
}

/* アクションセットをアンロード / Unload the action set */
function unloadForceNewGraphicStyleAction() {
    try {
        app.unloadAction(TEMP_STYLE_ACTION_SET, '');
    } catch (e) {
    }
}

/* 現在の選択を配列として退避 / Capture current selection as an array */
// app.selection が TextRange のときは .length が文字数を返すので、配列扱いせず単体で退避する
function captureCurrentDocumentSelection() {
    var items = [];
    try {
        var currentSelection = app.selection;
        if (!currentSelection) return items;
        if (currentSelection.typename === "TextRange") {
            items.push(currentSelection);
            return items;
        }
        if (typeof currentSelection.length === "number") {
            for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
                items.push(currentSelection[selectionIndex]);
            }
            return items;
        }
        items.push(currentSelection);
    } catch (e) {
    }
    return items;
}

/* 退避した選択を復元 / Restore captured selection */
// テキスト編集モードへの完全復帰はできないが、TextRange は単体で再選択する
function restoreDocumentSelection(items) {
    try {
        app.selection = null;
    } catch (clearError) {
    }
    if (!items || items.length === 0) return;

    if (items.length === 1 && items[0] && items[0].typename === "TextRange") {
        try {
            items[0].select();
            return;
        } catch (textRangeSelectError) {
        }
        try {
            app.selection = items[0];
            return;
        } catch (textRangeAssignError) {
        }
        return;
    }

    try {
        app.selection = items;
    } catch (assignError) {
        try {
            app.selection = items[0];
        } catch (singleAssignError) {
        }
    }
}

/* TextFrame の見た目を temp_style として登録 / Register the TextFrame's appearance as temp_style */
// register-temp-style.jsx と同じ手順：既存 temp_style を削除→アクションで末尾に追加→末尾を改名
function registerTextFrameAsTempGraphicStyle(textFrame) {
    if (!textFrame) return null;

    var activeDoc;
    try {
        activeDoc = app.activeDocument;
    } catch (e) {
        return null;
    }
    if (!activeDoc) return null;

    var graphicStyles;
    try {
        graphicStyles = activeDoc.graphicStyles;
    } catch (graphicStylesError) {
        return null;
    }
    if (!graphicStyles) return null;

    /* 既存の temp_style を削除 / Remove the existing temp_style */
    try {
        graphicStyles.getByName(TEMP_STYLE_NAME).remove();
    } catch (removeExistingError) {
    }

    var savedSelection = captureCurrentDocumentSelection();

    try {
        app.selection = null;
    } catch (clearSelectionError) {
    }

    try {
        textFrame.selected = true;
    } catch (selectError) {
        restoreDocumentSelection(savedSelection);
        return null;
    }

    var beforeCount;
    try {
        beforeCount = graphicStyles.length;
    } catch (countError) {
        restoreDocumentSelection(savedSelection);
        return null;
    }

    var actionSucceeded = false;
    try {
        loadForceNewGraphicStyleAction();
        runForceNewGraphicStyleAction();
        actionSucceeded = true;
    } catch (actionError) {
    }
    unloadForceNewGraphicStyleAction();

    var registeredName = null;
    if (actionSucceeded) {
        try {
            if (graphicStyles.length > beforeCount) {
                graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
                registeredName = TEMP_STYLE_NAME;
            }
        } catch (renameError) {
        }
    }

    restoreDocumentSelection(savedSelection);
    return registeredName;
}

/* 結果表示ダイアログ / Show result dialog */
function showResultDialog(info) {
    var dialog = new Window("dialog", L("copiedMessageTitle"));
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 16;

    function createInfoPanel(titleKey) {
        var panel = dialog.add("panel", undefined, L(titleKey));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [15, 20, 15, 10];
        panel.spacing = 8;
        panel.alignment = ["fill", "top"];
        return panel;
    }

    function addRow(parent, key, value) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = 8;
        var labelControl = row.add("statictext", undefined, labelText(key));
        labelControl.preferredSize.width = 180;
        labelControl.justify = "right";
        row.add("statictext", undefined, String(value));
    }

    var fontSizeLeadingPanel = createInfoPanel("fontSizeLeadingPanelTitle");
    var kerningPanel = createInfoPanel("kerningPanelTitle");
    var paragraphOtherPanel = createInfoPanel("paragraphOtherPanelTitle");
    var fillGraphicStylePanel = createInfoPanel("fillGraphicStylePanelTitle");

    addRow(fontSizeLeadingPanel, "postScriptName", info.font.name);
    addRow(fontSizeLeadingPanel, "fontFamilyName", info.font.family);
    addRow(fontSizeLeadingPanel, "fontStyle", info.font.style);
    addRow(fontSizeLeadingPanel, "fontSize", formatPointValueForDisplay(info.size));
    addRow(fontSizeLeadingPanel, "leading", info.autoLeading ? "—" : formatPointValueForDisplay(info.leading));
    addRow(fontSizeLeadingPanel, "autoLeading", formatBooleanLabel(info.autoLeading));

    addRow(kerningPanel, "kerningMethod", info.kerningMethodLabel);
    addRow(kerningPanel, "proportionalMetrics", formatBooleanLabel(info.proportionalMetrics));
    addRow(kerningPanel, "tracking", info.tracking);
    addRow(kerningPanel, "tsume", info.tsume);

    addRow(paragraphOtherPanel, "orientation", info.orientationLabel);
    addRow(paragraphOtherPanel, "justification", info.justificationLabel);

    /* 塗りとグラフィックスタイル：取得したものを並べて表示
       Fill & Graphic Style: show captured info side by side */
    addRow(fillGraphicStylePanel, "fillColor", info.fillColorLabel);
    addRow(fillGraphicStylePanel, "graphicStyle", info.graphicStyleName ? info.graphicStyleName : L("graphicStyleNotRegistered"));

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var closeButton = buttonGroup.add("button", undefined, L("closeButton"), { name: "ok" });
    closeButton.onClick = function () { dialog.close(); };

    dialog.show();
}

/* 行揃えの表示名を取得 / Get display label for paragraph justification */
function getJustificationLabel(justification) {
    try {
        if (justification === Justification.LEFT) {
            return (lang === "ja") ? "左揃え" : "Left";
        }
        if (justification === Justification.CENTER) {
            return (lang === "ja") ? "中央揃え" : "Center";
        }
        if (justification === Justification.RIGHT) {
            return (lang === "ja") ? "右揃え" : "Right";
        }
        if (justification === Justification.FULLJUSTIFYLASTLINELEFT) {
            return (lang === "ja") ? "均等配置（最終行左揃え）" : "Justify with last line aligned left";
        }
        if (justification === Justification.FULLJUSTIFYLASTLINECENTER) {
            return (lang === "ja") ? "均等配置（最終行中央揃え）" : "Justify with last line centered";
        }
        if (justification === Justification.FULLJUSTIFYLASTLINERIGHT) {
            return (lang === "ja") ? "均等配置（最終行右揃え）" : "Justify with last line aligned right";
        }
        if (justification === Justification.FULLJUSTIFY) {
            return (lang === "ja") ? "均等配置" : "Full justify";
        }
    } catch (e) {
        return "—";
    }
    return "—";
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

    var sourceTextFrame = getParentTextFrameFromTextRange(sourceTextRange);

    var copiedLeading = sourceCharacterAttributes.leading;
    var copiedAutoLeading = getAutoLeadingState(sourceTextRange, sourceFirstCharacter, sourceTextFrame);
    var copiedTsume = Math.round(sourceCharacterAttributes.Tsume);
    var copiedKerningMethod = sourceCharacterAttributes.kerningMethod;
    var copiedProportionalMetrics = sourceCharacterAttributes.proportionalMetrics;
    var copiedTracking = sourceCharacterAttributes.tracking;
    var copiedJustification = sourceParagraphAttributes.justification;
    var copiedJustificationLabel = getJustificationLabel(copiedJustification);

    var copiedOrientation = sourceTextFrame ? sourceTextFrame.orientation : null;
    var copiedOrientationLabel = getOrientationLabel(sourceTextFrame);
    var copiedKerningMethodLabel = getKerningMethodLabel(copiedKerningMethod);

    /* 先頭文字の塗り色を取得 / Get fill color from the first character */
    var copiedFillColor = cloneColor(getTextRangeFillColor(sourceFirstCharacter));
    var copiedFillColorLabel = formatColorForDisplay(copiedFillColor);

    /* 現在の TextFrame の見た目を temp_style として登録（既存があれば上書き）/
       Register the current TextFrame's appearance as temp_style (overwrites if exists) */
    var registeredGraphicStyleName = registerTextFrameAsTempGraphicStyle(sourceTextFrame);

    var hasUsableFillColor = !!(copiedFillColor && copiedFillColor.typename && copiedFillColor.typename !== "NoColor");

    $.global.FontClipboard = {
        font: {
            name: sourceCharacterAttributes.textFont.name,
            family: sourceCharacterAttributes.textFont.family,
            style: sourceCharacterAttributes.textFont.style
        },
        size: sourceCharacterAttributes.size,
        leading: copiedLeading,
        autoLeading: copiedAutoLeading,
        tsume: copiedTsume,
        tracking: copiedTracking,
        kerningMethod: copiedKerningMethod,
        kerningMethodLabel: copiedKerningMethodLabel,
        proportionalMetrics: copiedProportionalMetrics,
        orientation: copiedOrientation,
        orientationLabel: copiedOrientationLabel,
        justification: copiedJustification,
        justificationLabel: copiedJustificationLabel,
        fillColor: copiedFillColor,
        fillColorLabel: copiedFillColorLabel,
        /* 自動登録した temp_style の名前。Apply 側はこの名前でグラフィックスタイルを引く /
           Name of the auto-registered temp_style. The Apply script looks up the graphic style by this name. */
        graphicStyleName: registeredGraphicStyleName,
        /* Apply 側ラジオの初期選択：塗りがあれば fill、グラフィックスタイルだけあれば graphicStyle /
           Default radio in the Apply dialog: fill when a fill is captured, graphicStyle when only a style is, otherwise none */
        fillOrGraphicStyle: hasUsableFillColor ? "fill" : (registeredGraphicStyleName ? "graphicStyle" : "none")
        // 将来拡張
    };

    showResultDialog($.global.FontClipboard);
})();