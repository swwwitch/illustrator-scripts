#target illustrator
#targetengine "FontClipboard"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyTextAttributesFromClipboard
Meta-Default-language=ja
Name=保存した文字属性をテキストに適用
Name[en]=Apply saved text attributes to the selection
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.11
Release-Date=2021-04-10
Tags=style,スタイル
X-com.swwwitch-Script-Name=ApplyTextAttributesFromClipboard
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyTextAttributesFromClipboard.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyTextAttributesFromClipboard.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ApplyTextAttributesFromClipboard.jsx
Description-BEGIN
[ja]
CopyTextAttributesToClipboard.jsx が保存した文字属性を、選択中のテキストへ適用します。
適用する属性はダイアログのチェックボックスで選べ、文字ツールでの部分選択にも対応します。
[en]
Applies the text attributes saved by CopyTextAttributesToClipboard.jsx to the current text selection.
Which attributes are applied is chosen with checkboxes in the dialog, and partial selections made with the Type tool are supported.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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
        tipFillGraphicStyleNone: {
            ja: "塗りもグラフィックスタイルも適用しません。文字の属性だけを反映します。",
            en: "Applies neither the fill nor the graphic style. Only the character attributes are copied."
        },
        tipFill: {
            ja: "コピー元の塗りカラーを適用します。",
            en: "Applies the fill color taken from the source."
        },
        tipGraphicStyle: {
            ja: "コピー元に付いていたグラフィックスタイルを適用します。ドキュメントに同名のスタイルが無いと適用できません。",
            en: "Applies the graphic style the source carried. It needs a style of the same name in this document."
        },
        tipPreview: {
            ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
            en: "Shows the result on the canvas. Cancel restores the original state."
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
        fillGraphicStyleNoneOption: {
            ja: "しない",
            en: "None"
        },
        graphicStyleLabel: {
            ja: "グラフィックスタイル",
            en: "Graphic Style"
        },
        graphicStyleNotFound: {
            ja: "（未登録）",
            en: "(not in document)"
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
        tsumeLabel: {
            ja: "文字ツメ",
            en: "Tsume"
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
        fillColorLabel: {
            ja: "塗り",
            en: "Fill"
        },
        fillColorNone: {
            ja: "なし",
            en: "None"
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

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    /* 数値を表示用に丸める / Round number for display */
    function formatNumberForDisplay(value, digits) {
        var multiplier = Math.pow(10, digits);
        var rounded = Math.round(value * multiplier) / multiplier;
        return String(rounded);
    }

    /* pt値を現在の文字単位で表示 / Format point value using current text unit */
    function formatPointValueForDialog(pointValue) {
        var textUnit = getUnitInfo("text/units");
        var displayValue = ptToUnit(pointValue, textUnit);
        return formatNumberForDisplay(displayValue, 3) + " " + textUnit.label;
    }

    // =========================================
    // 塗り色ユーティリティ / Fill color utilities
    // =========================================

    /* 色を安全に複製 / Safely clone a color */
    // 復元・適用時の独立性を保つため、毎回新規インスタンスを返す
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

    /* 色を表示用文字列へ整形 / Format color for display */
    function formatColorForDialog(color) {
        if (!color || !color.typename) return getLabel("fillColorNone");
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
                return getLabel("fillColorNone");
            default:
                return color.typename;
        }
    }

    /* テキスト範囲の塗り色を取得 / Get fill color from text range */
    function getTextRangeFillColor(textRange) {
        try {
            return textRange.characterAttributes.fillColor;
        } catch (e) {
            return null;
        }
    }

    /* テキスト範囲へ塗り色を適用 / Apply fill color to a text range */
    // textRange.characterAttributes に一括代入すると例外が出ないまま無視されてブラックになることがあるため、
    // FillStrokeSwitcher と同じく文字単位の代入を優先し、文字列挙ができないときだけ範囲一括にフォールバックする
    function applyFillColorToTextRange(textRange, color) {
        if (!textRange || !color) return false;

        var characters = null;
        try {
            characters = textRange.characters;
        } catch (e1) {
            characters = null;
        }

        if (characters && characters.length > 0) {
            var appliedAny = false;
            for (var i = 0; i < characters.length; i++) {
                var clonedForChar = cloneColor(color);
                if (!clonedForChar) continue;
                try {
                    characters[i].characterAttributes.fillColor = clonedForChar;
                    appliedAny = true;
                } catch (charError) {
                }
            }
            if (appliedAny) return true;
        }

        /* 文字列挙ができない場合の最終フォールバック / Final fallback when characters cannot be enumerated */
        var clonedForRange = cloneColor(color);
        if (!clonedForRange) return false;
        try {
            textRange.characterAttributes.fillColor = clonedForRange;
            return true;
        } catch (e2) {
        }

        return false;
    }

    // =========================================
    // グラフィックスタイルユーティリティ / Graphic style utilities
    // =========================================

    /* ドキュメント内のグラフィックスタイルを名前で検索 / Find a graphic style by name in the active document */
    function findGraphicStyleByName(name) {
        if (!name) return null;
        try {
            if (!app.activeDocument || !app.activeDocument.graphicStyles) return null;
            for (var i = 0; i < app.activeDocument.graphicStyles.length; i++) {
                try {
                    if (app.activeDocument.graphicStyles[i].name === name) {
                        return app.activeDocument.graphicStyles[i];
                    }
                } catch (graphicStyleNameError) {
                }
            }
        } catch (e) {
        }
        return null;
    }

    /* 名前を指定してグラフィックスタイルを安全に適用 / Apply a graphic style by name safely */
    function applyGraphicStyleSafely(textFrame, graphicStyleName) {
        if (!textFrame || !graphicStyleName) return false;
        var graphicStyle = findGraphicStyleByName(graphicStyleName);
        if (!graphicStyle) return false;
        try {
            graphicStyle.applyTo(textFrame);
            return true;
        } catch (e) {
            return false;
        }
    }

    /* グラフィックスタイル適用前の TextFrame 外観属性を退避 / Capture TextFrame appearance affected by graphic style */
    // 完全な外観復元は不可能だが、塗り・線・不透明度・描画モードを戻すことで多くのケースを救う
    function captureTextFrameAppearance(textFrame) {
        if (!textFrame) return null;
        var captured = {};
        for (var i = 0; i < APPEARANCE_PROPERTIES.length; i++) {
            copyAppearanceProperty(captured, textFrame, APPEARANCE_PROPERTIES[i]);
        }
        return captured;
    }

    /* 退避・復元する TextFrame の外観プロパティ / TextFrame appearance properties that are saved and restored */
    var APPEARANCE_PROPERTIES = [
        { name: "fillColor", isColor: true },
        { name: "strokeColor", isColor: true },
        { name: "filled", isColor: false },
        { name: "stroked", isColor: false },
        { name: "strokeWidth", isColor: false },
        { name: "opacity", isColor: false },
        { name: "blendingMode", isColor: false }
    ];

    /* 退避した TextFrame 外観属性を復元 / Restore captured TextFrame appearance */
    function restoreTextFrameAppearance(textFrame, captured) {
        if (!textFrame || !captured) return;
        for (var i = 0; i < APPEARANCE_PROPERTIES.length; i++) {
            var property = APPEARANCE_PROPERTIES[i];
            if (typeof captured[property.name] === "undefined" || captured[property.name] === null) continue;
            copyAppearanceProperty(textFrame, captured, property);
        }
    }

    /**
     * 外観プロパティを1つ写す
     * 種類によっては持っていないプロパティがあり、読み書きで例外になるため1つずつ受け流す。
     * @param {object} targetObject - 写し先
     * @param {object} sourceObject - 写し元
     * @param {object} property - APPEARANCE_PROPERTIES の1項目
     * @returns {void}
     */
    function copyAppearanceProperty(targetObject, sourceObject, property) {
        try {
            targetObject[property.name] = property.isColor ?
                cloneColor(sourceObject[property.name]) :
                sourceObject[property.name];
        } catch (e) {}
    }

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

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

        /* 選択の最上位にあるテキストフレームだけを対象にする / Only text frames at the top level of the selection */
        var textFrames = collectSelectionTextFrames(selection, { enterGroups: false });
        for (var frameIndex = 0; frameIndex < textFrames.length; frameIndex++) {
            targets.push({
                textRange: textFrames[frameIndex].textRange,
                textFrame: textFrames[frameIndex]
            });
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
            tsume: characterAttributes.Tsume,
            tracking: characterAttributes.tracking,
            kerningMethod: characterAttributes.kerningMethod,
            proportionalMetrics: characterAttributes.proportionalMetrics,
            orientation: textFrame ? textFrame.orientation : null,
            justification: paragraphAttributes.justification,
            paragraphJustifications: captureParagraphJustifications(textRange),
            fillColor: cloneColor(getTextRangeFillColor(firstCharacter)),
            /* グラフィックスタイル適用は TextFrame 外観を書き換えるため、退避しておく
               Graphic style application rewrites the TextFrame appearance, so capture it */
            textFrameAppearance: captureTextFrameAppearance(textFrame)
        };
    }

    /* 実際に適用した属性だけを退避値へ戻す / Restore only applied attributes to captured values */
    // appliedState を受け取り、プレビューで実際に変更した項目だけ退避値へ戻す
    // 混在書式の未適用項目を代表値で上書きしないため、通常の完全復元とは分けて扱う
    // （特に自動カーニングは復元のためにダイナミックアクションを読み込む必要があるため、未適用なら呼び出さない）
    function restoreAppliedTextAttributes(textRange, textFrame, capturedAttributes, appliedState) {
        if (!capturedAttributes) return;
        if (!appliedState) return;

        var characterAttributes = textRange.characterAttributes;

        if (appliedState.applyFont) {
            try {
                characterAttributes.textFont = app.textFonts.getByName(capturedAttributes.fontName);
            } catch (e) {
                // 復元時にフォントが見つからない場合は無視 / Ignore if the font cannot be found while restoring
            }
        }

        if (appliedState.applySize) {
            characterAttributes.size = capturedAttributes.size;
        }

        if (appliedState.applyAutoLeading) {
            characterAttributes.autoLeading = capturedAttributes.autoLeading;
        }

        if (appliedState.applyLeading && !capturedAttributes.autoLeading) {
            characterAttributes.autoLeading = false;
            characterAttributes.leading = capturedAttributes.leading;
        }

        if (appliedState.applyTsume) {
            characterAttributes.Tsume = Math.round(capturedAttributes.tsume);
        }

        if (appliedState.applyTracking) {
            characterAttributes.tracking = capturedAttributes.tracking;
        }

        /* 自動カーニングは適用していないときは復元不要（ダイナミックアクションのロードを避ける） */
        /*   Skip kerning restore unless it was actually applied (avoid loading the dynamic action) */
        if (appliedState && appliedState.applyKerningMethod) {
            applyKerningMethodSafely(characterAttributes, capturedAttributes, textRange);
        }

        if (appliedState.applyProportionalMetrics) {
            characterAttributes.proportionalMetrics = capturedAttributes.proportionalMetrics;
        }

        if (appliedState.applyJustification) {
            if (capturedAttributes.paragraphJustifications && capturedAttributes.paragraphJustifications.length > 0) {
                restoreParagraphJustifications(textRange, textFrame, capturedAttributes.paragraphJustifications);
            }
            if (typeof capturedAttributes.justification !== "undefined" && capturedAttributes.justification !== null) {
                applyJustificationSafely(textRange, textFrame, capturedAttributes.justification);
            }
        }

        if (appliedState.applyOrientation && textFrame && capturedAttributes.orientation !== null && typeof capturedAttributes.orientation !== "undefined") {
            textFrame.orientation = capturedAttributes.orientation;
        }

        /* 適用していたモードに応じて塗り or グラフィックスタイル外観を戻す /
           Restore fill or graphic-style appearance based on the mode that was applied */
        if (appliedState.fillGraphicStyleMode === "fill" && capturedAttributes.fillColor) {
            applyFillColorToTextRange(textRange, capturedAttributes.fillColor);
        } else if (appliedState.fillGraphicStyleMode === "graphicStyle" && capturedAttributes.textFrameAppearance) {
            restoreTextFrameAppearance(textFrame, capturedAttributes.textFrameAppearance);
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

    /* 複数対象で実際に適用した属性だけを退避値へ戻す / Restore only applied attributes for multiple targets */
    function restoreAppliedTextAttributesForTargets(targets, capturedList, appliedState) {
        if (!targets || !capturedList) return;
        for (var targetIndex = 0; targetIndex < targets.length; targetIndex++) {
            restoreAppliedTextAttributes(targets[targetIndex].textRange, targets[targetIndex].textFrame, capturedList[targetIndex], appliedState);
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
                runKerningActionWithSelection(textRange, function () { runKerningActionByType('metrics'); });
            } catch (actionError) {
                // アクションでメトリクスを適用できない場合はスキップ / Skip if Metrics cannot be applied by action
            }
            return;
        }

        if (isOpticalKerningMethod(copiedAttributes)) {
            try {
                runKerningActionWithSelection(textRange, function () { runKerningActionByType('optical'); });
            } catch (actionError2) {
                // アクションでオプティカルを適用できない場合はスキップ / Skip if Optical cannot be applied by action
            }
            return;
        }

        if (isZeroKerningMethod(copiedAttributes)) {
            try {
                runKerningActionWithSelection(textRange, function () { runKerningActionByType('zero'); });
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

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /* 自動カーニング用ダイナミックアクションのキャッシュ / Cache for auto kerning dynamic actions */
    // 同じ種別のアクションを複数テキストへ繰り返し適用しても loadAction が一度で済むよう、現在ロード中の種別を保持する
    // #targetengine "FontClipboard" はエンジンが残り続けるため、main の開始時に必ず resetKerningActionCache() を呼んで初期化する
    var kerningActionCache = {
        setName: 'Kerning',
        loadedType: null,
        aiaBodies: {
            metrics: '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 7' + ' 4d657472696373' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1635019621' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 0' + ' }' + ' }' + '}',
            optical: '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 7' + ' 4f70746963616c' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1869638501' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 1' + ' }' + ' }' + '}',
            zero: '/version 3' + '/name [ 7' + ' 4b65726e696e67' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 8' + ' 4b65726e696e6730' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOCharacterPalette)' + ' /localizedName [ 6' + ' e69687e5ad97' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1801810542' + ' /showInPalette 4294967295' + ' /type (integer)' + ' /value 0' + ' }' + ' }' + '}'
        },
        actionNames: {
            metrics: 'Metrics',
            optical: 'Optical',
            zero: 'Kerning0'
        }
    };

    /* キャッシュ状態を初期化 / Reset cache state */
    // #targetengine の永続性により持ち越される loadedType を毎回クリアする
    function resetKerningActionCache() {
        kerningActionCache.loadedType = null;
    }

    /* 指定種別のダイナミックアクションを必要なときだけ読み込む / Lazy-load the kerning action of the given type */
    function ensureKerningActionLoaded(type) {
        if (kerningActionCache.loadedType === type) return;

        /* 別種別がロード中の場合は入れ替え / Swap if a different type is currently loaded */
        if (kerningActionCache.loadedType !== null) {
            unloadTemporaryActionSet(kerningActionCache.setName);
            kerningActionCache.loadedType = null;
        }

        var actionBody = kerningActionCache.aiaBodies[type];
        if (!actionBody) return;

        /* 読み込めなければ loadedType は null のまま（実行側でスキップ）/ On failure loadedType stays null and the caller skips */
        if (loadTemporaryActionSet(actionBody, kerningActionCache.setName)) {
            kerningActionCache.loadedType = type;
        }
    }

    /* キャッシュされたカーニングアクションを実行 / Run a cached kerning action */
    function runKerningActionByType(type) {
        ensureKerningActionLoaded(type);
        if (kerningActionCache.loadedType !== type) return;
        app.doScript(kerningActionCache.actionNames[type], kerningActionCache.setName, false);
    }

    /* 読み込んだカーニングアクションを破棄 / Unload the cached kerning action */
    function unloadCachedKerningAction() {
        if (kerningActionCache.loadedType === null) return;
        unloadTemporaryActionSet(kerningActionCache.setName);
        kerningActionCache.loadedType = null;
    }

    /* 対応する行揃え値か判定 / Check whether the justification value is supported */
    function isSupportedJustification(justification) {
        try {
            return justification === Justification.LEFT ||
                justification === Justification.CENTER ||
                justification === Justification.RIGHT ||
                justification === Justification.FULLJUSTIFYLASTLINELEFT ||
                justification === Justification.FULLJUSTIFYLASTLINECENTER ||
                justification === Justification.FULLJUSTIFYLASTLINERIGHT ||
                justification === Justification.FULLJUSTIFY;
        } catch (e) {
            return false;
        }
    }

    /* 行揃えを安全に適用 / Apply justification safely */
    function applyJustificationSafely(textRange, textFrame, justification) {
        if (!textRange) return false;
        if (typeof justification === "undefined" || justification === null) return false;
        if (!isSupportedJustification(justification)) return false;

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
                throw new Error(getLabel("errorFontNotFoundPrefix") + copiedAttributes.font.name + getLabel("errorFontNotFoundSuffix"));
            }
        }

        if (uiState.applySize) {
            characterAttributes.size = copiedAttributes.size;
        }

        if (uiState.applyAutoLeading) {
            characterAttributes.autoLeading = copiedAttributes.autoLeading;
        }

        if (uiState.applyLeading && !copiedAttributes.autoLeading) {
            characterAttributes.autoLeading = false;
            characterAttributes.leading = copiedAttributes.leading;
        }

        if (uiState.applyTsume) {
            characterAttributes.Tsume = Math.round(copiedAttributes.tsume);
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

        /* モードに応じて塗り or グラフィックスタイルを適用 /
           Apply either fill or graphic style depending on the selected mode */
        if (uiState.fillGraphicStyleMode === "fill" && copiedAttributes.fillColor) {
            applyFillColorToTextRange(textRange, copiedAttributes.fillColor);
        } else if (uiState.fillGraphicStyleMode === "graphicStyle" && textFrame && copiedAttributes.graphicStyleName) {
            applyGraphicStyleSafely(textFrame, copiedAttributes.graphicStyleName);
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
        var fillGraphicStyleMode = "none";
        if (ui.rbFill && ui.rbFill.value) {
            fillGraphicStyleMode = "fill";
        } else if (ui.rbGraphicStyle && ui.rbGraphicStyle.value) {
            fillGraphicStyleMode = "graphicStyle";
        }
        return {
            applyFont: ui.cbFont.value,
            applySize: ui.cbSize.value,
            applyLeading: ui.cbLeading.value,
            applyAutoLeading: ui.cbAutoLeading.value,
            applyTsume: ui.cbTsume.value,
            applyTracking: ui.cbTracking.value,
            applyKerningMethod: ui.cbKerningMethod.value,
            applyProportionalMetrics: ui.cbProportionalMetrics.value,
            applyOrientation: ui.cbOrientation.value,
            applyJustification: ui.cbJustification.value,
            fillGraphicStyleMode: fillGraphicStyleMode
        };
    }

    /* 適用対象があるか確認 / Check whether any item should be applied */
    function hasAnyApplyTarget(uiState) {
        return uiState.applyFont ||
            uiState.applySize ||
            uiState.applyLeading ||
            uiState.applyAutoLeading ||
            uiState.applyTsume ||
            uiState.applyTracking ||
            uiState.applyKerningMethod ||
            uiState.applyProportionalMetrics ||
            uiState.applyOrientation ||
            uiState.applyJustification ||
            (uiState.fillGraphicStyleMode && uiState.fillGraphicStyleMode !== "none");
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
            typeof copiedAttributes.justification !== "undefined" &&
            isSupportedJustification(copiedAttributes.justification)
        );
    }

    /* コピー済みの塗り色があるか確認 / Check whether copied fill color exists */
    function hasCopiedFillColor(copiedAttributes) {
        if (!copiedAttributes) return false;
        var color = copiedAttributes.fillColor;
        if (!color || !color.typename) return false;
        return color.typename !== "NoColor";
    }

    /* コピー済みのグラフィックスタイル名があるか確認 / Check whether a copied graphic style name exists */
    function hasCopiedGraphicStyleName(copiedAttributes) {
        return !!(
            copiedAttributes &&
            typeof copiedAttributes.graphicStyleName === "string" &&
            copiedAttributes.graphicStyleName.length > 0
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
        return value ? getLabel("onValue") : getLabel("offValue");
    }

    /* 組み方向表示を返す / Return orientation display text */
    function formatOrientationForDialog(orientation) {
        if (orientation === TextOrientation.VERTICAL) {
            return getLabel("verticalOrientation");
        }
        return getLabel("horizontalOrientation");
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

        var checkboxLabel = isAvailable ? valueText : getLabel("notStored");
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

        var checkboxLabel = isAvailable ? fontText : getLabel("notStored");
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

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = getVersionInfoValue("version");
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // =========================================
    // メイン処理 / Main process
    // =========================================

    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("errorNoDocument"));
            return;
        }

        /* 同じ #targetengine の $.global.FontClipboard から取得 / Read from $.global.FontClipboard in the same #targetengine */
        var copiedTextAttributes = $.global.FontClipboard;

        var hasFont = hasCopiedFont(copiedTextAttributes);
        var hasSize = hasCopiedSize(copiedTextAttributes);
        var hasLeading = hasCopiedNumber(copiedTextAttributes, "leading");
        var hasAutoLeading = hasCopiedBoolean(copiedTextAttributes, "autoLeading");
        var hasTsume = hasCopiedNumber(copiedTextAttributes, "tsume");
        var hasTracking = hasCopiedNumber(copiedTextAttributes, "tracking");
        var hasKerningMethod = hasCopiedKerningMethod(copiedTextAttributes);
        var hasProportionalMetrics = hasCopiedBoolean(copiedTextAttributes, "proportionalMetrics");
        var hasOrientation = hasCopiedOrientation(copiedTextAttributes);
        var hasJustification = hasCopiedJustification(copiedTextAttributes);
        var hasFillColor = hasCopiedFillColor(copiedTextAttributes);
        var hasGraphicStyleNameStored = hasCopiedGraphicStyleName(copiedTextAttributes);

        if (!hasFont && !hasSize && !hasLeading && !hasAutoLeading && !hasTsume && !hasTracking && !hasKerningMethod && !hasProportionalMetrics && !hasOrientation && !hasJustification && !hasFillColor && !hasGraphicStyleNameStored) {
            alert(getLabel("errorNoCopiedAttributes"));
            return;
        }

        var applyTargets = getApplyTargetsFromSelection(app.selection);

        if (applyTargets.length === 0) {
            alert(getLabel("errorNoTextSelection"));
            return;
        }

        for (var applyTargetIndex = 0; applyTargetIndex < applyTargets.length; applyTargetIndex++) {
            if (applyTargets[applyTargetIndex].textRange.characters.length === 0) {
                alert(getLabel("errorEmptyTextRange"));
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
        var dlg = new Window("dialog", getLabel("dialogTitle"));
        setupWindow(dlg);

        var fontDisplay = hasFont ? (copiedTextAttributes.font.family || getLabel("notStored")) : getLabel("notStored");
        var fontStyleDisplay = hasFont ? (copiedTextAttributes.font.style || getLabel("notStored")) : getLabel("notStored");
        var sizeDisplay = hasSize ? formatPointValueForDialog(copiedTextAttributes.size) : getLabel("notStored");
        var leadingDisplay = hasLeading ? formatPointValueForDialog(copiedTextAttributes.leading) : getLabel("notStored");
        var autoLeadingDisplay = hasAutoLeading ? formatBooleanForDialog(copiedTextAttributes.autoLeading) : getLabel("notStored");
        var tsumeDisplay = hasTsume ? String(copiedTextAttributes.tsume) : getLabel("notStored");
        var trackingDisplay = hasTracking ? String(copiedTextAttributes.tracking) : getLabel("notStored");
        var kerningMethodDisplay = hasKerningMethod
            ? (copiedTextAttributes.kerningMethodLabel || String(copiedTextAttributes.kerningMethod))
            : getLabel("notStored");
        var proportionalMetricsDisplay = hasProportionalMetrics ? formatBooleanForDialog(copiedTextAttributes.proportionalMetrics) : getLabel("notStored");
        var orientationDisplay = hasOrientation
            ? (copiedTextAttributes.orientationLabel || formatOrientationForDialog(copiedTextAttributes.orientation))
            : getLabel("notStored");
        var justificationDisplay = hasJustification
            ? (copiedTextAttributes.justificationLabel || String(copiedTextAttributes.justification))
            : getLabel("notStored");
        var fillColorDisplay = hasFillColor
            ? (copiedTextAttributes.fillColorLabel || formatColorForDialog(copiedTextAttributes.fillColor))
            : getLabel("notStored");

        var fontSizeLeadingPanel = dlg.add("panel", undefined, getLabel("fontSizeLeadingPanelTitle"));
        setupPanel(fontSizeLeadingPanel, 6);
        fontSizeLeadingPanel.alignChildren = "left";

        var cbFont = addFontCheckboxRow(fontSizeLeadingPanel, fontDisplay, fontStyleDisplay, hasFont, true, ATTRIBUTE_LABEL_WIDTH);
        var cbSize = addAttributeCheckboxRow(fontSizeLeadingPanel, "sizeLabel", sizeDisplay, hasSize, false, ATTRIBUTE_LABEL_WIDTH);
        var cbLeading = addAttributeCheckboxRow(fontSizeLeadingPanel, "leadingLabel", leadingDisplay, hasLeading, false, ATTRIBUTE_LABEL_WIDTH);
        /* コピー元が自動行送りの場合は、行送りの適用はできないためチェックボックスを無効化
           Disable leading checkbox when the copied source uses auto leading */
        if (hasAutoLeading && copiedTextAttributes.autoLeading === true) {
            cbLeading.enabled = false;
            cbLeading.value = false;
        }
        var cbAutoLeading = addAttributeCheckboxRow(fontSizeLeadingPanel, "autoLeadingLabel", autoLeadingDisplay, hasAutoLeading, false, ATTRIBUTE_LABEL_WIDTH);

        var kerningPanel = dlg.add("panel", undefined, getLabel("kerningPanelTitle"));
        setupPanel(kerningPanel, 6);
        kerningPanel.alignChildren = "left";

        var cbKerningMethod = addAttributeCheckboxRow(kerningPanel, "kerningMethodLabel", kerningMethodDisplay, hasKerningMethod, false, ATTRIBUTE_LABEL_WIDTH);
        var cbProportionalMetrics = addAttributeCheckboxRow(kerningPanel, "proportionalMetricsLabel", proportionalMetricsDisplay, hasProportionalMetrics, false, ATTRIBUTE_LABEL_WIDTH);
        var cbTracking = addAttributeCheckboxRow(kerningPanel, "trackingLabel", trackingDisplay, hasTracking, false, ATTRIBUTE_LABEL_WIDTH);
        var cbTsume = addAttributeCheckboxRow(kerningPanel, "tsumeLabel", tsumeDisplay, hasTsume, false, ATTRIBUTE_LABEL_WIDTH);

        var paragraphOtherPanel = dlg.add("panel", undefined, getLabel("paragraphOtherPanelTitle"));
        setupPanel(paragraphOtherPanel, 6);
        paragraphOtherPanel.alignChildren = "left";

        var cbOrientation = addAttributeCheckboxRow(paragraphOtherPanel, "orientationLabel", orientationDisplay, hasOrientation, false, ATTRIBUTE_LABEL_WIDTH);
        cbOrientation.enabled = hasOrientation && canApplyOrientation;
        var cbJustification = addAttributeCheckboxRow(paragraphOtherPanel, "justificationLabel", justificationDisplay, hasJustification, false, ATTRIBUTE_LABEL_WIDTH);

        /* 塗りとグラフィックスタイル：3択ラジオで排他 / Fill & Graphic Style: 3-way radio */
        var fillGraphicStylePanel = dlg.add("panel", undefined, getLabel("fillGraphicStylePanelTitle"));
        setupPanel(fillGraphicStylePanel, 6);
        fillGraphicStylePanel.alignChildren = "left";

        var hasCopiedGSName = hasCopiedGraphicStyleName(copiedTextAttributes);
        var graphicStyleExistsInDoc = hasCopiedGSName && !!findGraphicStyleByName(copiedTextAttributes.graphicStyleName);
        var graphicStyleRowText;
        if (hasCopiedGSName) {
            graphicStyleRowText = graphicStyleExistsInDoc
                ? copiedTextAttributes.graphicStyleName
                : copiedTextAttributes.graphicStyleName + " " + getLabel("graphicStyleNotFound");
        } else {
            graphicStyleRowText = getLabel("notStored");
        }

        var FILL_GS_RADIO_NAME_WIDTH = 165;

        /* 「しない」行 / "None" row */
        var fgsNoneRow = fillGraphicStylePanel.add("group");
        fgsNoneRow.orientation = "row";
        fgsNoneRow.alignChildren = ["left", "center"];
        fgsNoneRow.spacing = 6;
        var rbFillGraphicStyleNone = fgsNoneRow.add("radiobutton", undefined, getLabel("fillGraphicStyleNoneOption"));
        rbFillGraphicStyleNone.helpTip = getLabel("tipFillGraphicStyleNone");

        /* 「塗り」行 / "Fill" row */
        var fgsFillRow = fillGraphicStylePanel.add("group");
        fgsFillRow.orientation = "row";
        fgsFillRow.alignChildren = ["left", "center"];
        fgsFillRow.spacing = 6;
        var rbFill = fgsFillRow.add("radiobutton", undefined, getLabel("fillColorLabel"));
        rbFill.helpTip = getLabel("tipFill");
        rbFill.preferredSize.width = FILL_GS_RADIO_NAME_WIDTH;
        rbFill.enabled = hasFillColor;
        fgsFillRow.add("statictext", undefined, hasFillColor ? fillColorDisplay : getLabel("notStored"));

        /* 「グラフィックスタイル」行 / "Graphic Style" row */
        var fgsGSRow = fillGraphicStylePanel.add("group");
        fgsGSRow.orientation = "row";
        fgsGSRow.alignChildren = ["left", "center"];
        fgsGSRow.spacing = 6;
        var rbGraphicStyle = fgsGSRow.add("radiobutton", undefined, getLabel("graphicStyleLabel"));
        rbGraphicStyle.helpTip = getLabel("tipGraphicStyle");
        rbGraphicStyle.preferredSize.width = FILL_GS_RADIO_NAME_WIDTH;
        rbGraphicStyle.enabled = graphicStyleExistsInDoc && canApplyOrientation;
        fgsGSRow.add("statictext", undefined, graphicStyleRowText);

        /* 初期選択：Copy 側で保存された fillOrGraphicStyle に従う /
           Default selection: follow the fillOrGraphicStyle saved by the Copy script */
        var copiedFillGSMode = copiedTextAttributes && copiedTextAttributes.fillOrGraphicStyle;
        if (copiedFillGSMode === "graphicStyle" && rbGraphicStyle.enabled) {
            rbGraphicStyle.value = true;
        } else if (copiedFillGSMode === "fill" && rbFill.enabled) {
            rbFill.value = true;
        } else {
            rbFillGraphicStyleNone.value = true;
        }

        /* 別グループに置いたラジオは自動排他にならないので手動で同期する /
           Radios in separate groups are not auto-exclusive in ScriptUI, so sync manually */
        var fillGraphicStyleRadios = [rbFillGraphicStyleNone, rbFill, rbGraphicStyle];
        function enforceFillGraphicStyleExclusivity(selected) {
            for (var fgsIndex = 0; fgsIndex < fillGraphicStyleRadios.length; fgsIndex++) {
                fillGraphicStyleRadios[fgsIndex].value = (fillGraphicStyleRadios[fgsIndex] === selected);
            }
        }

        var attributeCheckboxes = [
            cbFont,
            cbSize,
            cbLeading,
            cbAutoLeading,
            cbKerningMethod,
            cbProportionalMetrics,
            cbTracking,
            cbTsume,
            cbOrientation,
            cbJustification
        ];
        bindExclusiveOptionClick(attributeCheckboxes);

        /* ボタンエリア（cbPreviewを含む）を先に生成し、updatePreviewから安全に参照できるようにする

           Build the button area (including cbPreview) first so updatePreview can reference it safely */

        var buttonRow = addButtonRow(dlg);
        var cbPreview = buttonRow.leftGroup.add("checkbox", undefined, getLabel("previewCheckbox"));
        cbPreview.helpTip = getLabel("tipPreview");
        cbPreview.value = false;

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("cancelButton"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("applyButton"), { name: "ok" });

        /* スクリプト実行ごとにアクションキャッシュを初期化（#targetengine の永続性対策）
           Reset the action cache at the start of each run (handles #targetengine persistence) */
        resetKerningActionCache();

        var originalTextAttributesList = captureCurrentTextAttributesForTargets(applyTargets);
        var isPreviewApplied = false;
        /* 直近の適用状態を保持し、復元時に「実際に適用した項目だけ戻す」判定に使う
           Track the last applied state so restore can selectively undo only what was applied */
        var lastAppliedState = null;

        try {
            function removePreview() {
                if (!isPreviewApplied) return;
                restoreAppliedTextAttributesForTargets(applyTargets, originalTextAttributesList, lastAppliedState);
                isPreviewApplied = false;
                lastAppliedState = null;
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
                        cbTsume: cbTsume,
                        cbTracking: cbTracking,
                        cbKerningMethod: cbKerningMethod,
                        cbProportionalMetrics: cbProportionalMetrics,
                        cbOrientation: cbOrientation,
                        cbJustification: cbJustification,
                        rbFill: rbFill,
                        rbGraphicStyle: rbGraphicStyle
                    });
                    applyCopiedTextAttributesToTargets(applyTargets, copiedTextAttributes, previewState);
                    isPreviewApplied = true;
                    lastAppliedState = previewState;
                    app.redraw();
                } catch (e) {
                    restoreAppliedTextAttributesForTargets(applyTargets, originalTextAttributesList, lastAppliedState);
                    isPreviewApplied = false;
                    lastAppliedState = null;
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

            /* 塗り/グラフィックスタイルのラジオ：排他化＋プレビュー更新 /
               Fill/Graphic-style radios: enforce exclusivity and trigger preview */
            for (var fgsRadioIndex = 0; fgsRadioIndex < fillGraphicStyleRadios.length; fgsRadioIndex++) {
                (function (radioButton) {
                    radioButton.onClick = function () {
                        enforceFillGraphicStyleExclusivity(radioButton);
                        updatePreview();
                    };
                })(fillGraphicStyleRadios[fgsRadioIndex]);
            }

            alignRightOnlyButtonRow(buttonRow);
            addVersionFooter(dlg);
            prepareDialogWindow(dlg, SCRIPT_NAME);
            var dialogResult = dlg.show();
            if (dialogResult !== 1) {
                restoreAppliedTextAttributesForTargets(applyTargets, originalTextAttributesList, lastAppliedState);
                isPreviewApplied = false;
                lastAppliedState = null;
                return;
            }

            var applyState = readApplyUIState({
                cbFont: cbFont,
                cbSize: cbSize,
                cbLeading: cbLeading,
                cbAutoLeading: cbAutoLeading,
                cbTsume: cbTsume,
                cbTracking: cbTracking,
                cbKerningMethod: cbKerningMethod,
                cbProportionalMetrics: cbProportionalMetrics,
                cbOrientation: cbOrientation,
                cbJustification: cbJustification,
                rbFill: rbFill,
                rbGraphicStyle: rbGraphicStyle
            });

            if (!hasAnyApplyTarget(applyState)) {
                restoreAppliedTextAttributesForTargets(applyTargets, originalTextAttributesList, lastAppliedState);
                isPreviewApplied = false;
                lastAppliedState = null;
                alert(getLabel("errorNoApplyItem"));
                return;
            }

            /* 適用 / Apply */
            // textRange 全体に対して characterAttributes を書き換える / Rewrite characterAttributes for the entire textRange
            // プレビューONで表示済みの状態は適用内容と一致するため、戻す→再適用はスキップ /
            // When preview is ON, the on-screen state already matches applyState, so skip the revert/re-apply cycle
            try {
                if (isPreviewApplied && cbPreview.value) {
                    isPreviewApplied = false;
                    /* lastAppliedState は表示中の状態を保持し続ける / Keep lastAppliedState reflecting the current on-screen state */
                } else {
                    removePreview();
                    applyCopiedTextAttributesToTargets(applyTargets, copiedTextAttributes, applyState);
                    lastAppliedState = applyState;
                }
            } catch (e) {
                restoreAppliedTextAttributesForTargets(applyTargets, originalTextAttributesList, lastAppliedState);
                isPreviewApplied = false;
                lastAppliedState = null;
                alert(e.message);
                return;
            }
        } finally {
            /* スクリプト終了時にダイナミックアクションをアンロード / Unload the dynamic action when the script ends */
            unloadCachedKerningAction();
        }
    })();

})();
