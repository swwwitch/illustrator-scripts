#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

var SCRIPT_VERSION = "v1.6";

/*
 * FontCatalogGenerator for Adobe Illustrator
 *
 * [概要]
 * 更新日: 2026-01-27
 * システムにインストールされているフォントを一覧化し、
 * アートボード上にフォント見本を自動生成するためのスクリプトです。
 *
 * 表示文字列とフォントサイズはダイアログで指定できます。
 * ドキュメント上でテキストを選択している場合、その文字列とフォントを初期値として利用できます。
 *
 * 以下の条件でフォントを自動的に除外できます：
 * ・除外リスト内のキーワードを含むフォント（永続保存対応）
 * ・斜体フォント（Italic / Oblique）※ダイアログでON/OFF可能
 *
 * また、選択中テキストのフォントを
 * 除外リストに追加するUIを備えています。
 *
 * フォントサイズとアートボードサイズから
 * 1列あたりの行数を自動計算し、左端に10ptの余白を設けて配置します。
 *
 * 処理中はプログレスバーで進捗を表示します。
 *
 * [更新履歴]
 * 2026-01-27
 * - v1.1: Error 9（予約語 'var'）の原因だった配列内 var 宣言を修正（systemFontKeywords と illustratorBundledFontKeywords を分離）
 * - v1.2: 除外判定の精度改善（system/lang/variable を name+family+style の大文字小文字無視検索に統一し、空白差異などによる除外漏れを低減）
 * - v1.3: 中国語（繁体）除外キーワードに「標楷體」系を追加（標楷體-港澳/標楷體-繁 などが除外漏れしていたため）
 * - v1.4: AdobeCleanUX が除外されていなかったため、システムフォント除外キーワードに追加
 * - v1.5: ヘブライ語／タイ語／アラビア系フォントの除外オプションを追加（fontSearchText の大文字小文字無視マッチで判定）
 * - v1.6: 永続除外リストの照合を強化（name+family+style の大文字小文字無視＋空白差異の吸収）— Font Awesome などの除外漏れ対策
 * - ダイアログUIを追加（表示文字／フォントサイズ）
 * - 選択中テキストの内容を初期値に流用
 * - 斜体（Italic / Oblique）除外をダイアログで制御可能に
 * - 除外リストを Preferences（JSON）に永続保存
 * - 選択フォントを除外リストに追加するUIを追加
 * - フォントサイズとアートボードサイズから行数を自動計算
 * - プログレスバーを追加
 *
 * オリジナル：studio TOFU
 * https://note.com/studio_tofu/n/n7b0cf367ec88
 * 
 */

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: {
    ja: "フォント見本作成",
    en: "FontCatalogGenerator"
  },
  labelSampleText: {
    ja: "表示する文字:",
    en: "Sample text:"
  },
  labelFontSize: {
    ja: "フォントサイズ:",
    en: "Font size:"
  },
  panelSampleSettings: {
    ja: "見本の設定",
    en: "Sample settings"
  },
  panelFontNameDisplay: {
    ja: "フォント名の表示",
    en: "Font name display"
  },
  panelExcludeList: {
    ja: "除外リスト",
    en: "Exclude list"
  },
  panelExcludeAdd: {
    ja: "除外リストへの追加",
    en: "Add to exclude list"
  },
  checkboxExcludeItalic: {
    ja: "斜体（Italic/Oblique）",
    en: "Italic/Oblique"
  },
  checkboxExcludeKorean: {
    ja: "韓国語フォント",
    en: "Korean fonts"
  },
  checkboxExcludeChineseSC: {
    ja: "中国語フォント（簡体）",
    en: "Chinese fonts (Simplified)"
  },
  checkboxExcludeChineseTC: {
    ja: "中国語フォント（繁体）",
    en: "Chinese fonts (Traditional)"
  },
  checkboxExcludeHebrew: {
    ja: "ヘブライ語フォント",
    en: "Hebrew fonts"
  },
  checkboxExcludeThai: {
    ja: "タイ語フォント",
    en: "Thai fonts"
  },
  checkboxExcludeArabic: {
    ja: "アラビア系フォント",
    en: "Arabic fonts"
  },
  checkboxExcludeSystemFonts: {
    ja: "システムフォント",
    en: "System fonts"
  },
  checkboxExcludeVariableFonts: {
    ja: "バリアブル",
    en: "Variable fonts"
  },
  checkboxExcludeIllustratorBundled: {
    ja: "Illustrator付属",
    en: "Bundled with Illustrator"
  },
  checkboxExcludeCompositeFonts: {
    ja: "合成フォント",
    en: "Composite fonts"
  },
  checkboxExcludeMorisawa: {
    ja: "モリサワ",
    en: "Morisawa"
  },
  checkboxExcludeFontworks: {
    ja: "フォントワークス",
    en: "Fontworks"
  },
  checkboxAddSelectedFontToExclude: {
    ja: "選択フォントを除外",
    en: "Exclude selected font"
  },
  labelSelectedFontPrefix: {
    ja: "選択フォント: ",
    en: "Selected font: "
  },
  labelSelectedFontNone: {
    ja: "（なし）",
    en: "(none)"
  },
  buttonCancel: {
    ja: "キャンセル",
    en: "Cancel"
  },
  buttonOK: {
    ja: "OK",
    en: "OK"
  },
  progressTitle: {
    ja: "処理中...",
    en: "Processing..."
  },
  progressCancel: {
    ja: "キャンセル",
    en: "Cancel"
  },
  progressCancelled: {
    ja: "キャンセルしました",
    en: "Cancelled"
  },
  checkboxShowFontName: {
    ja: "フォント名（ファミリー＋スタイル）",
    en: "Font name (family + style)"
  },
  checkboxShowPostScriptName: {
    ja: "PostScript名（内部名）",
    en: "PostScript name (internal)"
  }
};


function L(key) {
  try {
    return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
  } catch (e) {
    return key;
  }
}

/* --------------------------------------------------
 * 単位ユーティリティ（文字単位対応） / Unit utilities for text
 * -------------------------------------------------- */

// 単位コード → ラベル（簡易）
var unitLabelMap = {
  0: "in",
  1: "mm",
  2: "pt",
  3: "pica",
  4: "cm",
  5: "Q/H",
  6: "px",
  7: "ft/in",
  8: "m",
  9: "yd",
  10: "ft"
};

// 現在の文字単位ラベルを取得 / Get current text unit label
function getCurrentTextUnitLabel() {
  try {
    var unitCode = app.preferences.getIntegerPreference("text/units");
    return unitLabelMap[unitCode] || "pt";
  } catch (e) {
    return "pt";
  }
}

// 単位コード → pt 換算係数 / Convert unit code to pt factor
function getPtFactorFromUnitCode(code) {
  switch (code) {
    case 0: return 72.0;                        // in
    case 1: return 72.0 / 25.4;                 // mm
    case 2: return 1.0;                         // pt
    case 3: return 12.0;                        // pica
    case 4: return 72.0 / 2.54;                 // cm
    case 5: return 72.0 / 25.4 * 0.25;          // Q or H
    case 6: return 1.0;                         // px
    case 7: return 72.0 * 12.0;                 // ft/in
    case 8: return 72.0 / 25.4 * 1000.0;        // m
    case 9: return 72.0 * 36.0;                 // yd
    case 10:return 72.0 * 12.0;                 // ft
    default:return 1.0;
  }
}

// pt → 現在の文字単位値 / Convert pt to current text unit
function ptToCurrentTextUnit(ptValue) {
  try {
    var code = app.preferences.getIntegerPreference("text/units");
    var factor = getPtFactorFromUnitCode(code);
    return ptValue / factor;
  } catch (e) {
    return ptValue;
  }
}

// 現在の文字単位値 → pt / Convert current text unit value to pt
function currentTextUnitToPt(unitValue) {
  try {
    var code = app.preferences.getIntegerPreference("text/units");
    var factor = getPtFactorFromUnitCode(code);
    return unitValue * factor;
  } catch (e) {
    return unitValue;
  }
}

// 表示用の数値整形 / Format number for display (trim trailing zeros)
function formatNumberForDisplay(num) {
  try {
    var n = Math.round(num * 1000) / 1000; // up to 3 decimals
    var s = String(n);
    if (s.indexOf(".") !== -1) {
      s = s.replace(/\.?0+$/, "");
    }
    return s;
  } catch (e) {
    return String(num);
  }
}

/* 部分一致（大文字小文字無視） / Case-insensitive contains */
function containsIgnoreCase(haystack, needle) {
  try {
    if (!haystack || !needle) return false;
    return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
  } catch (e) {
    return false;
  }
}

/* 値の変更（↑↓キー対応） / Change value by arrow keys */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減 / Increment by 0.1 with Option
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め / Round to 1 decimal place
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め / Round to integer
            value = Math.round(value);
        }

        editText.text = value;
    });
}

(function() {
    // --- 除外リスト保存ファイル / Exclude list persistence ---
    var EXCLUDE_LIST_FILE = File("~/Library/Preferences/FontCatalogExcludeList.json");

    function loadExcludeList(defaultList) {
        try {
            if (EXCLUDE_LIST_FILE.exists) {
                EXCLUDE_LIST_FILE.open("r");
                var json = EXCLUDE_LIST_FILE.read();
                EXCLUDE_LIST_FILE.close();
                var data = JSON.parse(json);
                if (data && data instanceof Array) {
                    return data;
                }
            }
        } catch (e) {}
        return defaultList.slice(); // fallback (clone)
    }

    function saveExcludeList(list) {
        try {
            EXCLUDE_LIST_FILE.open("w");
            EXCLUDE_LIST_FILE.write(JSON.stringify(list, null, 2));
            EXCLUDE_LIST_FILE.close();
        } catch (e) {}
    }

    // --- 1. 基本設定 ---
    var defaultText = "Sample / サンプル"; 
    var defaultFontSizePt = 24; // internal pt

    // 選択中のテキストを初期値に流用（可能な場合） / Use selected text as default when possible
    // 併せて、選択中テキストのフォント名を取得（可能な場合） / Also capture selected text font name
    var hasSelectedText = false;
    var selectedFontName = "";
    try {
        if (app.documents.length > 0) {
            var selectionItems = app.activeDocument.selection;
            if (selectionItems && selectionItems.length > 0) {
                var firstSelectionItem = selectionItems[0];

                // TextFrame
                if (firstSelectionItem.typename === "TextFrame") {
                    hasSelectedText = true;
                    if (firstSelectionItem.contents) defaultText = firstSelectionItem.contents;
                    try {
                        selectedFontName = firstSelectionItem.textRange.characters[0].characterAttributes.textFont.name;
                    } catch (e1) { selectedFontName = ""; }

                // TextRange
                } else if (firstSelectionItem.typename === "TextRange") {
                    hasSelectedText = true;
                    if (firstSelectionItem.contents) defaultText = firstSelectionItem.contents;
                    try {
                        selectedFontName = firstSelectionItem.characters[0].characterAttributes.textFont.name;
                    } catch (e2) { selectedFontName = ""; }
                }
            }
        }
    } catch (e) {}

    // --- ダイアログボックス / Dialog ---
    var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    // --- 2カラムレイアウト / Two-column layout ---
    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = ["fill", "top"];

    var leftColumn = columnsGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["fill", "top"];

    var rightColumn = columnsGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["fill", "top"];

    // --- 表示 / Display ---
    var displayPanel = leftColumn.add("panel", undefined, L("panelSampleSettings"));
    displayPanel.orientation = "column";
    displayPanel.alignChildren = ["fill", "top"];
    displayPanel.margins = [15, 20, 15, 10];

    var textInputGroup = displayPanel.add("group");
    textInputGroup.orientation = "column";
    textInputGroup.alignChildren = ["fill", "top"];
    var sampleTextInput = textInputGroup.add("edittext", undefined, defaultText);
    sampleTextInput.characters = 20;

    var fontSizeGroup = displayPanel.add("group");
    fontSizeGroup.orientation = "row";

    var textUnitLabel = getCurrentTextUnitLabel();

    // ラベル末尾の「:」を除去（ja/en両対応） / Remove trailing colon from label
    var fontSizeLabelText = String(L("labelFontSize")).replace(/[:：]\s*$/, "");
    fontSizeGroup.add("statictext", undefined, fontSizeLabelText);

    var defaultFontSizeDisplay = ptToCurrentTextUnit(defaultFontSizePt);
    var fontSizeInput = fontSizeGroup.add("edittext", undefined, formatNumberForDisplay(defaultFontSizeDisplay));
    fontSizeInput.characters = 3;
    changeValueByArrowKey(fontSizeInput);

    fontSizeGroup.add("statictext", undefined, "(" + textUnitLabel + ")");

    // --- 表示オプション / Display options ---
    var displayOptionGroup = rightColumn.add("panel", undefined, L("panelFontNameDisplay"));
    displayOptionGroup.orientation = "column";
    displayOptionGroup.alignChildren = ["left", "top"];
    displayOptionGroup.margins = [15, 20, 15, 10];

    var showFontNameCheckbox = displayOptionGroup.add(
        "checkbox",
        undefined,
        L("checkboxShowFontName")
    );
    showFontNameCheckbox.value = true;

    var showPostScriptNameCheckbox = displayOptionGroup.add(
        "checkbox",
        undefined,
        L("checkboxShowPostScriptName")
    );
    showPostScriptNameCheckbox.value = false;

    // --- 全幅（貫通）: 除外リスト / Full width: Exclude list ---
    var fullWidthGroup = dialog.add("group");
    fullWidthGroup.orientation = "column";
    fullWidthGroup.alignChildren = ["fill", "top"];

    var excludeListPanel = fullWidthGroup.add("panel", undefined, L("panelExcludeList"));
    excludeListPanel.orientation = "column";
    excludeListPanel.alignChildren = ["fill", "top"];
    excludeListPanel.margins = [15, 20, 15, 10];

    // --- 除外オプション（3カラム） / Exclusion options (3 columns) ---
    var excludeColumnsGroup = excludeListPanel.add("group");
    excludeColumnsGroup.orientation = "row";
    excludeColumnsGroup.alignChildren = ["left", "top"];

    // 1列目 / Column 1
    var excludeCol1 = excludeColumnsGroup.add("group");
    excludeCol1.orientation = "column";
    excludeCol1.alignChildren = ["left", "top"];

    var excludeItalicCheckbox = excludeCol1.add(
        "checkbox",
        undefined,
        L("checkboxExcludeItalic")
    );
    excludeItalicCheckbox.value = true; // 既定で除外 / Default: exclude

    var excludeSystemFontsCheckbox = excludeCol1.add(
        "checkbox",
        undefined,
        L("checkboxExcludeSystemFonts")
    );
    excludeSystemFontsCheckbox.value = true;

    var excludeVariableFontsCheckbox = excludeCol1.add(
        "checkbox",
        undefined,
        L("checkboxExcludeVariableFonts")
    );
    excludeVariableFontsCheckbox.value = true;

    var excludeIllustratorBundledCheckbox = excludeCol1.add(
        "checkbox",
        undefined,
        L("checkboxExcludeIllustratorBundled")
    );
    excludeIllustratorBundledCheckbox.value = true;

    var excludeCompositeFontsCheckbox = excludeCol1.add(
        "checkbox",
        undefined,
        L("checkboxExcludeCompositeFonts")
    );
    excludeCompositeFontsCheckbox.value = true;

    // 2列目 / Column 2
    var excludeCol2 = excludeColumnsGroup.add("group");
    excludeCol2.orientation = "column";
    excludeCol2.alignChildren = ["left", "top"];

    var excludeKoreanCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeKorean")
    );
    excludeKoreanCheckbox.value = true;

    var excludeChineseSCCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeChineseSC")
    );
    excludeChineseSCCheckbox.value = true;

    var excludeChineseTCCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeChineseTC")
    );
    excludeChineseTCCheckbox.value = true;

    var excludeHebrewCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeHebrew")
    );
    excludeHebrewCheckbox.value = true;

    var excludeThaiCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeThai")
    );
    excludeThaiCheckbox.value = true;

    var excludeArabicCheckbox = excludeCol2.add(
        "checkbox",
        undefined,
        L("checkboxExcludeArabic")
    );
    excludeArabicCheckbox.value = true;

    // 3列目 / Column 3
    var excludeCol3 = excludeColumnsGroup.add("group");
    excludeCol3.orientation = "column";
    excludeCol3.alignChildren = ["left", "top"];

    var excludeMorisawaCheckbox = excludeCol3.add(
        "checkbox",
        undefined,
        L("checkboxExcludeMorisawa")
    );
    excludeMorisawaCheckbox.value = true;

    var excludeFontworksCheckbox = excludeCol3.add(
        "checkbox",
        undefined,
        L("checkboxExcludeFontworks")
    );
    excludeFontworksCheckbox.value = true;

    // --- 除外リスト / Exclude list ---
    var excludeOptionGroup = excludeCol3.add("panel", undefined, L("panelExcludeAdd"));
    excludeOptionGroup.orientation = "column";
    excludeOptionGroup.alignChildren = ["left", "top"];
    excludeOptionGroup.margins = [15, 20, 15, 10];

    var addSelectedFontToExcludeCheckbox = excludeOptionGroup.add(
        "checkbox",
        undefined,
        L("checkboxAddSelectedFontToExclude")
    );
    addSelectedFontToExcludeCheckbox.value = false;

    var selectedFontLabelText = hasSelectedText && selectedFontName
        ? (L("labelSelectedFontPrefix") + selectedFontName)
        : (L("labelSelectedFontPrefix") + L("labelSelectedFontNone"));
    var selectedFontLabel = excludeOptionGroup.add("statictext", undefined, selectedFontLabelText);

    // 選択テキストがない場合は無効化 / Disable when there is no selected text font
    if (!(hasSelectedText && selectedFontName)) {
        addSelectedFontToExcludeCheckbox.enabled = false;
    }


    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var cancelButton = buttonGroup.add("button", undefined, L("buttonCancel"), {name:"cancel"});
    var okButton = buttonGroup.add("button", undefined, L("buttonOK"), {name:"ok"});

    if (dialog.show() !== 1) return;

    var inputStr = sampleTextInput.text;
    var fontSizeInputValue = Number(fontSizeInput.text); // value in current text unit
    var fontSize = currentTextUnitToPt(fontSizeInputValue); // internal pt

    var showFontName = showFontNameCheckbox.value === true;
    var showPostScriptName = showPostScriptNameCheckbox.value === true;

    var excludeItalicFonts = excludeItalicCheckbox.value === true;
    var excludeKoreanFonts = excludeKoreanCheckbox.value === true;
    var excludeChineseSCFonts = excludeChineseSCCheckbox.value === true;
    var excludeChineseTCFonts = excludeChineseTCCheckbox.value === true;
    var excludeHebrewFonts = excludeHebrewCheckbox.value === true;
    var excludeThaiFonts = excludeThaiCheckbox.value === true;
    var excludeArabicFonts = excludeArabicCheckbox.value === true;
    var excludeSystemFonts = excludeSystemFontsCheckbox.value === true;
    var excludeVariableFonts = excludeVariableFontsCheckbox.value === true;
    var excludeIllustratorBundledFonts = excludeIllustratorBundledCheckbox.value === true;
    var excludeCompositeFonts = excludeCompositeFontsCheckbox.value === true;
    var excludeMorisawaFonts = excludeMorisawaCheckbox.value === true;
    var excludeFontworksFonts = excludeFontworksCheckbox.value === true;
    var addSelectedFontToExclude = addSelectedFontToExcludeCheckbox.value === true;
    // バリアブルフォント / Variable fonts
    var variableFontKeywords = [
        "Variable",
        "VF",
        "Var"
    ];
    if (!inputStr || isNaN(fontSizeInputValue) || fontSizeInputValue <= 0) return;
    if (isNaN(fontSize) || fontSize <= 0) return;

    var columnGap = 80;   // 列同士の間隔
    var rowGap = 20;      // 行同士の間隔
    var rowLimit = 20;    // 1列に並べる個数（後で自動計算で上書き） / Will be overwritten by auto calc

    var infoFontSize = 9;        // 下部表示テキストのサイズ / Info text size
    var infoTopGap = 2;          // 見本と情報の間隔 / Gap between sample and info
    var infoBottomGap = 6;       // 情報と次行の間隔 / Gap after info block

    var infoFontPostScriptName = "HiraginoSans-W3"; // 情報表示用フォント（PostScript名） / Info font (PostScript)
    var infoTextFont = null;
    try {
        infoTextFont = app.textFonts.getByName(infoFontPostScriptName);
    } catch (e) {
        infoTextFont = null;
    }

    var documentRef = app.documents.add();
    var allFonts = app.textFonts;
    var fontCount = allFonts.length; // パフォーマンス改善のためキャッシュ
    
    // --- 進捗表示（プログレスバー） / Progress bar ---
    var progressPalette = new Window("palette", L("progressTitle"));
    progressPalette.orientation = "column";
    progressPalette.alignChildren = ["fill", "top"];

    var progressLabel = progressPalette.add("statictext", undefined, "0 / " + fontCount);

    var progressBar = progressPalette.add("progressbar", undefined, 0, fontCount);
    progressBar.preferredSize.width = 320;

    var progressButtonGroup = progressPalette.add("group");
    progressButtonGroup.alignment = "right";

    var progressCancelButton = progressButtonGroup.add("button", undefined, L("progressCancel"));
    var isCancelled = false;

    progressCancelButton.onClick = function () {
        isCancelled = true;
        try {
            progressCancelButton.enabled = false;
            progressLabel.text = L("progressCancelled");
            progressPalette.update();
        } catch (e) {}
    };

    progressPalette.show();

    // アートボードの座標取得 [左, 上, 右, 下]
    var activeArtboard = documentRef.artboards[documentRef.artboards.getActiveArtboardIndex()];
    var artboardRect = activeArtboard.artboardRect; 

    // 開始座標を左上に設定（左に10ptの余白、1行目がはみ出ないようfontSize分だけ下にオフセット）
    var leftMarginPt = 10; // pt
    var startPosX = artboardRect[0] + leftMarginPt; 
    var startPosY = artboardRect[1] - fontSize; 

    // rowLimit をフォントサイズとアートボード高さから自動計算 / Auto-calc rowLimit from font size and artboard height
    // 1行目は startPosY、以降は (fontSize + rowGap) ずつ下げるため、入る行数 = floor(availableHeight / step) + 1
    try {
        var topY = startPosY;
        var bottomY = artboardRect[3];
        var availableHeight = topY - bottomY;
        var stepY = fontSize + rowGap;
        if (stepY > 0) {
            var autoRowLimit = Math.floor(availableHeight / stepY) + 1;
            if (autoRowLimit < 1) autoRowLimit = 1;
            rowLimit = autoRowLimit;
        }
    } catch (e) {}

    var currentX = startPosX;
    var currentY = startPosY;
    var maxColumnWidth = 0;
    var placedItemCount = 0; 

    // 言語別キーワード（UIでON/OFF制御） / Language keyword groups (controlled by UI)
    var koreanFontKeywords = [
        "Korean",
        "Hangul",
        "AppleSDGothicNeo",
        "Malgun",
        "Gulim",
        "Dotum",
        "Batang",
        "Nanum",
        "AdobeMyungjoStd",    // Adobe 명조 Std (Korean)
        "AdobeGothicStd"     // Adobe 고딕 Std (Korean)
    ];

    // 中国語（簡体） / Chinese (Simplified)
    var chineseSCFontKeywords = [
        "Chinese",
        "PingFangSC",
        "PingFang SC",
        "HeitiSC",
        "Heiti SC",
        "SongtiSC",
        "Songti SC",
        "SimSun",
        "SimHei",
        "MicrosoftYaHei",
        "Microsoft YaHei",
        "GB",
        "AdobeHeitiStd",      // Adobe 黑体
        "AdobeKaitiStd",      // Adobe 楷体
        "AdobeFangsongStd",   // Adobe 仿宋
        "AdobeSongStd",       // Adobe 宋体
        "文鼎",
        "ARS",
        "Arphic",
        "AR PL"
    ];

    // 中国語（繁体） / Chinese (Traditional)
    var chineseTCFontKeywords = [
        "Chinese",
        "標楷體",
        "標楷體-港澳",
        "標楷體-繁",
        "PingFangTC",
        "PingFang TC",
        "HeitiTC",
        "Heiti TC",
        "SongtiTC",
        "Songti TC",
        "MingLiU",
        "PMingLiU",
        "DFKai",
        "Big5",
        "AdobeMingStd",       // Adobe 明體 Std (Traditional Chinese)
        "AdobeHeitiStd",      // Adobe 黑体
        "AdobeKaitiStd",      // Adobe 楷体
        "AdobeFangsongStd",   // Adobe 仿宋
        "AdobeSongStd",       // Adobe 宋体
        "文鼎",
        "ARS",
        "Arphic",
        "AR PL"
    ];

    // ヘブライ語 / Hebrew
    var hebrewFontKeywords = [
        "Hebrew",
        "MyriadHebrew",
        "AdobeHebrew",
        "Narkisim",
        "Ezra",
        "FrankRuehl",
        "David"
    ];

    // タイ語 / Thai
    var thaiFontKeywords = [
        "Thai",
        "Thonburi",
        "Ayuthaya",
        "Krungthep",
        "Sukhumvit",
        "Silom",
        "NotoSansThai",
        "NotoSerifThai"
    ];

    // アラビア系 / Arabic
    var arabicFontKeywords = [
        "Arabic",
        "AdobeArabic",
        "Geeza",
        "Diwan",
        "Kufi",
        "Naskh",
        "NotoNaskhArabic",
        "NotoKufiArabic",
        "NotoSansArabic"
    ];

    // メーカー別キーワード（UIでON/OFF制御） / Vendor keyword groups (controlled by UI)
    var morisawaFontKeywords = [
        "A-OTF",
        "A P-OTF",
        "AP-OTF",
        "A-SK"
    ];

    var fontworksFontKeywords = [
        "FOT-",   // explicit prefix used by many Fontworks PS names
        "FOT"     // fallback
    ];

    // システムフォント（macOS） / System fonts (macOS)
    var systemFontKeywords = [
        "AdobeCleanUX",
        "ADTNumeric",
        "Apple Braille",
        "Apple Symbols",
        "AppleSDGothicNeo",
        "AquaKana",
        "ArialHB",
        "Avenir",
        "Avenir Next",
        "Avenir Next Condensed",
        "CJKSymbolsFallback",
        "Courier",
        "DecoTypeNastaleeqUrdu",
        "GeezaPro",
        "Geneva",
        "HelveLTMM",
        "Helvetica",
        "HelveticaNeue",
        "Hiragino Sans GB",
        "Keyboard",
        "Kohinoor",
        "KohinoorBangla",
        "KohinoorGujarati",
        "KohinoorTelugu",
        "LastResort",
        "LucidaGrande",
        "MarkerFelt",
        "Menlo",
        "Monaco",
        "MuktaMahee",
        "NewYork",
        "Noteworthy",
        "Optima",
        "Palatino",
        "SF",
        "STHeiti",
        "Symbol",
        "ThonburiUI",
        "Times",
        "TimesLTMM",
        "ZapfDingbats",
        "Zither",
        "ヒラギノ角ゴシック",
        "ヒラギノ丸ゴ",
        "ヒラギノ明朝"
    ];

    // Illustrator付属フォント / Bundled with Illustrator
    // Note: list is based on provided file names; we match by partial keyword against fontItem.name/family/style.
    var illustratorBundledFontKeywords = [
        "AcuminVariableConcept",
        "AdobeArabic",
        "AdobeInvisible",
        "AdobeMingStd",
        "AdobeMyungjoStd",
        "AdobeSongStd",
        "EmojiOneColor",
        "KozGoPr6N",
        "MinionVariableConcept",
        "MyriadHebrew",
        "MyriadPro",
        "MyriadVariableConcept",
        "SourceCodeVariable",
        "SourceHanSansSC",
        "SourceSansVariable",
        "SourceSerifVariable",
        "TrajanColor"
    ];

    // 合成フォント（ATC-） / Composite fonts (ATC-*)
    var compositeFontPrefix = "ATC-";

    function isKeywordInList(keyword, list) {
        for (var ii = 0; ii < list.length; ii++) {
            if (list[ii] === keyword) return true;
        }
        return false;
    }

    function containsAnyKeyword(text, list) {
        for (var kk = 0; kk < list.length; kk++) {
            if (text.indexOf(list[kk]) !== -1) return true;
        }
        return false;
    }

    // 複数フィールドをまとめてキーワード検索（大文字小文字無視） / Keyword search across fields (case-insensitive)
    function containsAnyKeywordCI(text, list) {
        for (var kk = 0; kk < list.length; kk++) {
            if (containsIgnoreCase(text, list[kk])) return true;
        }
        return false;
    }

    // --- 2. 除外リスト (編集エリア) ---
    // 非表示にしたいフォントのキーワードを自由に足してください。
    // より具体的なパターンを先に配置
    var excludeList = loadExcludeList([
        "HiraMin",           // ヒラギノ明朝
        "HiraKaku",          // ヒラギノ角ゴ
        "Hiragino",          // ヒラギノシリーズ
        "MS-P",              // MS Pゴシック/明朝
        "MS-Gothic",         // MS ゴシック
        "MS-Mincho",         // MS 明朝
        "YuMincho",          // 游明朝
        "YuGothic",          // 游ゴシック
        "Meiryo",            // メイリオ
        "Noto",              // Noto フォント全般
        "NotoSans",          // Noto Sans
        "NotoSerif",         // Noto Serif
        "SourceHan",         // 源ノ角ゴシック/源ノ明朝
        "Koz",               // 小塚ゴシック/明朝
        "AdobeSongStd",      // Adobe 宋体
        "BIZ-UD",            // BIZ UDシリーズ
        "UDDigiKyo",         // UD デジタル教科書体
        "HGP",               // HG系 (P)
        "HGS",               // HG系 (S)
        "HG",                // HG系
        "Font Awesome",      // Font Awesome (icon font)
        "Adobe Clean",       // Adobe Clean (UI/system font)
        "Artifakt Element",  // Artifakt Element (UI font)
        "Ornaments",         // Ornaments (icon/ornament fonts)
        // STIX math fonts
        "STIXIntegralsD",
        "STIXIntegralsUpD",
        "STIXSizeFiveSym",
        "STIXSizeThreeSym",
        "STIXIntegralsSm",
        "STIXIntegralsUpSm",
        "STIXSizeFourSym",
        "STIXSizeTwoSym",
        "STIXIntegralsUp",
        "STIXNonUnicode",
        "STIXSizeOneSym",
        "STIXVariants",
        "AppleColorEmoji",   // 絵文字 (Mac)
        "SegoeUIEmoji"       // 絵文字 (Win)
    ]);

    // 韓国語/中国語/システムフォントキーワードは専用チェックで制御するため、永続リストからは除外 / Keep language/system keywords out of persisted list
    try {
        var cleaned = [];
        for (var c = 0; c < excludeList.length; c++) {
            var kw2 = excludeList[c];
            if (isKeywordInList(kw2, koreanFontKeywords)) continue;
            if (isKeywordInList(kw2, chineseSCFontKeywords)) continue;
            if (isKeywordInList(kw2, chineseTCFontKeywords)) continue;
            if (isKeywordInList(kw2, hebrewFontKeywords)) continue;
            if (isKeywordInList(kw2, thaiFontKeywords)) continue;
            if (isKeywordInList(kw2, arabicFontKeywords)) continue;
            if (isKeywordInList(kw2, systemFontKeywords)) continue;
            if (isKeywordInList(kw2, variableFontKeywords)) continue;
            if (isKeywordInList(kw2, illustratorBundledFontKeywords)) continue;
            if (kw2.indexOf("ATC-") === 0) continue;
            cleaned.push(kw2);
        }
        excludeList = cleaned;
    } catch (eClean) {}

    // 選択フォントを除外リストに追加（任意） / Optionally add selected font to exclude list
    if (addSelectedFontToExclude && selectedFontName) {
        var alreadyExists = false;
        for (var k = 0; k < excludeList.length; k++) {
            if (excludeList[k] === selectedFontName) {
                alreadyExists = true;
                break;
            }
        }
        if (!alreadyExists) {
            excludeList.unshift(selectedFontName); // 先頭に追加（優先） / Add to top for priority
        }
    }
    // 除外リストを保存 / Save exclude list
    saveExcludeList(excludeList);

    // --- 3. メイン処理 ---
    for (var i = 0; i < fontCount; i++) {
        // progress update
        try {
            progressBar.value = i;
            progressLabel.text = i + " / " + fontCount;
            progressPalette.update();
        } catch (e) {}
        // cancel check
        if (isCancelled) {
            break;
        }

        var textFrame = null; // テキストフレームの参照を保持 / Keep refs for cleanup
        var infoTextFrame = null; // フォント情報表示用 / For font info display
        try {
            var fontItem = allFonts[i];
            if (!fontItem) continue;
            
            var fontName = fontItem.name;
            var fontStyle = "";
            try { fontStyle = fontItem.style; } catch(e) { fontStyle = ""; }

            // 斜体を除外（任意） / Optionally exclude italic fonts
            if (excludeItalicFonts) {
                if (fontName.indexOf("Italic") !== -1 || 
                    fontName.indexOf("Oblique") !== -1 || 
                    fontStyle.indexOf("Italic") !== -1 ||
                    fontStyle.indexOf("Oblique") !== -1) {
                    continue;
                }
            }

            // メーカー別フォント除外（任意） / Optionally exclude vendor fonts
            // PostScript名だけでなく family/style も含めて判定 / Check name + family/style for robustness
            var fontSearchText = fontName;
            try {
                var fam2 = "";
                var sty2 = "";
                try { fam2 = fontItem.family || ""; } catch (eFam2) { fam2 = ""; }
                try { sty2 = fontItem.style || ""; } catch (eSty2) { sty2 = ""; }
                fontSearchText = fontName + " | " + fam2 + " | " + sty2;
            } catch (eFS) {}

            if (excludeMorisawaFonts) {
                if (containsAnyKeywordCI(fontSearchText, morisawaFontKeywords)) {
                    continue;
                }
            }
            if (excludeFontworksFonts) {
                if (containsAnyKeywordCI(fontSearchText, fontworksFontKeywords)) {
                    continue;
                }
            }

            // システムフォント除外（任意） / Optionally exclude system fonts
            // name/family/style をまとめて大文字小文字無視で判定（空白差異なども吸収） / Case-insensitive across name/family/style
            if (excludeSystemFonts) {
                if (containsAnyKeywordCI(fontSearchText, systemFontKeywords)) {
                    continue;
                }
            }

            // バリアブルフォント除外（任意） / Optionally exclude variable fonts
            if (excludeVariableFonts) {
                if (containsAnyKeywordCI(fontSearchText, variableFontKeywords)) {
                    continue;
                }
            }

            // Illustrator付属フォント除外（任意） / Optionally exclude Illustrator bundled fonts
            if (excludeIllustratorBundledFonts) {
                if (containsAnyKeywordCI(fontSearchText, illustratorBundledFontKeywords)) {
                    continue;
                }
            }

            // 合成フォント除外（任意） / Optionally exclude composite fonts (ATC-*)
            if (excludeCompositeFonts) {
                if (fontName.indexOf(compositeFontPrefix) === 0) {
                    continue;
                }
            }

            // 言語別フォント除外（任意） / Optionally exclude language fonts
            if (excludeKoreanFonts) {
                if (containsAnyKeywordCI(fontSearchText, koreanFontKeywords)) {
                    continue;
                }
            }
            if (excludeChineseSCFonts) {
                if (containsAnyKeywordCI(fontSearchText, chineseSCFontKeywords)) {
                    continue;
                }
            }
            if (excludeChineseTCFonts) {
                if (containsAnyKeywordCI(fontSearchText, chineseTCFontKeywords)) {
                    continue;
                }
            }
            if (excludeHebrewFonts) {
                if (containsAnyKeywordCI(fontSearchText, hebrewFontKeywords)) {
                    continue;
                }
            }
            if (excludeThaiFonts) {
                if (containsAnyKeywordCI(fontSearchText, thaiFontKeywords)) {
                    continue;
                }
            }
            if (excludeArabicFonts) {
                if (containsAnyKeywordCI(fontSearchText, arabicFontKeywords)) {
                    continue;
                }
            }

            // 除外リスト照合（言語グループはUIでON/OFF） / Exclude list match (language groups are toggleable)
            // name+family+style で大文字小文字無視、さらに空白差異も吸収 / Case-insensitive across name/family/style + ignore whitespace differences
            var isFontExcluded = false;
            var hay = fontSearchText || fontName || "";
            var hayNoSpace = String(hay).replace(/\s+/g, "");

            for (var j = 0; j < excludeList.length; j++) {
                var kw = excludeList[j];
                if (!kw) continue;

                // 1) 通常（大文字小文字無視） / Normal case-insensitive match
                if (containsIgnoreCase(hay, kw)) {
                    isFontExcluded = true;
                    break;
                }

                // 2) 空白除去して比較（"Font Awesome" vs "FontAwesome" 等） / Ignore whitespace differences
                var kwNoSpace = String(kw).replace(/\s+/g, "");
                if (kwNoSpace && containsIgnoreCase(hayNoSpace, kwNoSpace)) {
                    isFontExcluded = true;
                    break;
                }
            }
            if (isFontExcluded) continue;

            // 描画テスト（合成フォント・エラーフォント回避）
            textFrame = documentRef.textFrames.add();
            textFrame.contents = inputStr;
            
            try {
                textFrame.textRange.characterAttributes.textFont = fontItem;
                textFrame.textRange.characterAttributes.size = fontSize;
            } catch(e) {
                if (textFrame && textFrame.parent) textFrame.remove();
                if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
                continue;
            }

            // フォント適用確認（フォールバック検知）
            if (textFrame.textRange.characters[0].characterAttributes.textFont.name !== fontName) {
                if (textFrame && textFrame.parent) textFrame.remove();
                if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
                continue;
            }

            // --- 配置 / Placement ---
            textFrame.left = currentX;
            textFrame.top = currentY;

            // フォント情報を見本の下に表示（任意） / Optionally show font info under the sample
            if (showFontName || showPostScriptName) {
                var fontObj = textFrame.textRange.characterAttributes.textFont;
                var family = "";
                var style = "";
                try { family = fontObj.family; } catch (eFam) { family = ""; }
                try { style = fontObj.style; } catch (eSty) { style = ""; }

                var postScriptName = "";
                try { postScriptName = textFrame.textRange.characterAttributes.textFont.name; } catch (ePS) { postScriptName = ""; }

                var nameWithStyle = family;
                if (style) {
                    nameWithStyle += " " + style;
                }

                var fontLine = "";

                if (showFontName && showPostScriptName) {
                    // 2 lines:
                    // <family> <style>
                    // <PostScriptName>
                    fontLine = nameWithStyle + "\r" + postScriptName;
                } else if (showFontName) {
                    // 1 line: <family> <style>
                    fontLine = nameWithStyle;
                } else if (showPostScriptName) {
                    // 1 line: <PostScriptName>
                    fontLine = postScriptName;
                }

                infoTextFrame = documentRef.textFrames.add();
                infoTextFrame.contents = fontLine;

                // プロポーショナルメトリクスをON（情報表示のみ） / Enable proportional metrics for info text only
                try { infoTextFrame.textRange.proportionalMetrics = true; } catch (ePM) {}

                if (infoTextFont) {
                    try { infoTextFrame.textRange.characterAttributes.textFont = infoTextFont; } catch (eFont) {}
                }
                infoTextFrame.textRange.characterAttributes.size = infoFontSize;

                // 情報は見本の下に配置 / Place info under the sample
                infoTextFrame.left = currentX;
                infoTextFrame.top = currentY - (fontSize + infoTopGap);

                // 列幅計算に反映 / Include in column width
                if (infoTextFrame.width > maxColumnWidth) {
                    maxColumnWidth = infoTextFrame.width;
                }

                // 次の行へ（情報ブロック分も下げる） / Move down including info block height
                currentY -= (fontSize + rowGap + infoTextFrame.height + infoTopGap + infoBottomGap);
            } else {
                // 次の行へ / Next row
                currentY -= (fontSize + rowGap);
            }

            if (textFrame.width > maxColumnWidth) {
                maxColumnWidth = textFrame.width;
            }

            placedItemCount++;

            // 20個並んだら次の列へ
            if (placedItemCount % rowLimit === 0) {
                currentX += (maxColumnWidth + columnGap);
                currentY = startPosY;
                maxColumnWidth = 0;
            }

        } catch (err) {
            // エラー時はテキストフレームを確実に削除
            try {
                if (textFrame && textFrame.parent) textFrame.remove();
                if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
            } catch(e) {}
        }
    }
    // 完了時にプログレスバーを閉じる
    try {
        progressBar.value = fontCount;
        progressLabel.text = fontCount + " / " + fontCount;
        progressPalette.update();
        progressPalette.close();
    } catch (e) {}
})();