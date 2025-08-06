#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

TypefaceSampler.jsx

### 概要

- Illustrator で利用可能なフォントを一覧表示し、ウエイトやスタイル順にアートボード上へ整列描画するスクリプトです。
- フォント名、PostScript名、サンプルテキスト、カスタムテキストなど柔軟に出力内容を切り替えられます。

### 主な機能

- ウエイト・スタイル評価によるスコアリングとソート
- フォント名、PostScript名、アルファベット、数字、カスタムテキスト表示に対応
- キーワード検索（AND/OR/NOT/先頭一致）によるフォント絞り込み
- カテゴリ（family）単位のグループ化と表示
- ウエイト数やスコアの表示切り替え
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログで出力形式やキーワード、オプションを設定
2. 検索条件に合致するフォントを収集
3. スコア順にソートし、アートボードに整列描画

### 更新履歴

- v1.0.0 (20250420) : 初期バージョン
- v1.3.0 (20250508) : 評価ロジック修正、条件絞り込み機能強化
- v1.3.1 (20250706) : ローカライズ調整

---

### Script Name:

TypefaceSampler.jsx

### Overview

- A script for Illustrator to list available fonts and draw them on the artboard sorted by weight and style.
- Supports flexible output options such as font name, PostScript name, sample text, and custom text.

### Main Features

- Weight/style score evaluation and sorting
- Supports display of font name, PostScript name, alphabet, numbers, or custom text
- Keyword search with AND/OR/NOT/prefix matching for font filtering
- Grouping by category (family) and display
- Toggle display of weight count and score
- Japanese and English UI support

### Process Flow

1. Configure output format, keywords, and options in dialog
2. Collect fonts matching search criteria
3. Sort by score and draw aligned on artboard

### Update History

- v1.0.0 (20250420): Initial version
- v1.3.0 (20250508): Improved evaluation logic and filter features
- v1.3.1 (20250706): Localization adjustments
*/


// -------------------------------
// 言語設定の取得
// Get current language setting
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// ラベル定義
// Define labels
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: {
        ja: "フォントを一覧表示",
        en: "Typeface Sampler"
    },
    keywordPrompt: {
        ja: "フォント名に含まれるキーワード（空欄→全対象）：",
        en: "Keyword in font name (leave blank for all):"
    },
    outputContent: {
        ja: "出力内容",
        en: "Output Content"
    },
    fontNameWeightStyle: {
        ja: "フォント名＋ウェイト／スタイル",
        en: "Font Name + Weight/Style"
    },
    postscriptName: {
        ja: "PostScript名",
        en: "PostScript Name"
    },
    sampleAlphabet: {
        ja: "The quick brown fox jumps over the lazy dog.",
        en: "The quick brown fox jumps over the lazy dog."
    },
    sampleNumbers: {
        ja: "1234567890",
        en: "1234567890"
    },
    custom: {
        ja: "カスタム",
        en: "Custom"
    },
    sampleText: {
        ja: "愛のあるユニークで豊かな書体ABCabcGg349",
        en: "Lorem ipsum dolor sit amet, consectetur adipiscing elit"
    },
    columns: {
        ja: "列数",
        en: "Columns"
    },
    categorizeByFamily: {
        ja: "カテゴリー分け（font.family）",
        en: "Categorize by font.family"
    },
    showWeightCount: {
        ja: "ウェイト数",
        en: "Weight Count"
    },
    showWeightToggle: {
        ja: "ウェイト一覧",
        en: "Weight List"
    },
    showScoreToggle: {
        ja: "スコア（検証用）",
        en: "Debug Score"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    confirmTitle: {
        ja: "確認",
        en: "Confirmation"
    },
    confirmMessage1: {
        ja: "すべてのフォントを対象に実行しますか？",
        en: "Do you want to process all fonts?"
    },
    confirmMessage2: {
        ja: "非常に時間がかかることがあります。",
        en: "This may take a long time."
    },
    stop: {
        ja: "中止する",
        en: "Cancel"
    },
    proceed: {
        ja: "続行する",
        en: "Proceed"
    },
    errorNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errorOccurred: {
        ja: "エラーが発生しました：",
        en: "An error occurred:"
    },
    labelGroupTitle: {
        ja: "[カテゴリ]",
        en: "[Category]"
    }
};
// -------------------------------------------
// ラジオボタンで上下キー移動を有効化する関数（オリジナル版）
// -------------------------------------------
function enableArrowKeyNavigation(radioButtons) {
    if (!radioButtons || radioButtons.length === 0) return;

    var parent = radioButtons[0].parent;
    while (parent && typeof parent.addEventListener !== "function" && parent.parent) {
        parent = parent.parent;
    }
    if (!parent || typeof parent.addEventListener !== "function") return;

    parent.addEventListener("keydown", function(event) {
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var currentIndex = -1;
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                currentIndex = i;
                break;
            }
        }

        if (currentIndex === -1) return;

        var nextIndex = currentIndex;
        if (event.keyName === "Up") {
            nextIndex = (currentIndex === 0) ? radioButtons.length - 1 : currentIndex - 1;
        } else if (event.keyName === "Down") {
            nextIndex = (currentIndex === radioButtons.length - 1) ? 0 : currentIndex + 1;
        }

        radioButtons[nextIndex].value = true;
        radioButtons[nextIndex].active = true;
        if (typeof radioButtons[nextIndex].onClick === "function") {
            radioButtons[nextIndex].onClick();
        }
        event.preventDefault && event.preventDefault();
    });
}

// --------------------------------
// ダイアログを表示してユーザー入力を取得
// Show dialog to get user input
// --------------------------------
function showFontListDialog() {
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";
    dialog.margins = 20;

    dialog.add("statictext", undefined, LABELS.keywordPrompt[lang]);
    var keywordField = dialog.add("edittext", undefined, "");
    keywordField.characters = 30;
    keywordField.active = true;

    var displayPanel = dialog.add("panel", undefined, LABELS.outputContent[lang]);
    displayPanel.orientation = "column";
    displayPanel.alignChildren = "left";
    displayPanel.margins = [15, 20, 15, 15];

    var radioGroup = displayPanel.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var displayOptions = [];
    // Only 5 elements, starting with "fontNameWeightStyle"
    displayOptions[0] = radioGroup.add("radiobutton", undefined, LABELS.fontNameWeightStyle[lang]);
    displayOptions[1] = radioGroup.add("radiobutton", undefined, LABELS.postscriptName[lang]);
    displayOptions[2] = radioGroup.add("radiobutton", undefined, LABELS.sampleAlphabet[lang]);
    displayOptions[3] = radioGroup.add("radiobutton", undefined, LABELS.sampleNumbers[lang]);
    displayOptions[4] = radioGroup.add("radiobutton", undefined, LABELS.custom[lang]);
    displayOptions[0].value = true;

    var customTextField = displayPanel.add("edittext", undefined, LABELS.sampleText[lang]);
    customTextField.characters = 30;
    customTextField.enabled = false;

    displayOptions[4].onClick = function() {
        customTextField.enabled = true;
        updateDisplayPanelState();
    };
    for (var i = 0; i < displayOptions.length - 1; i++) {
        displayOptions[i].onClick = function() {
            customTextField.enabled = false;
            updateDisplayPanelState();
        };
    }

    // --- Enable arrow key navigation for radio buttons ---
    enableArrowKeyNavigation(displayOptions);

    // 新しいパネル「表示オプション」
    var optionPanel = dialog.add("panel", undefined, LABELS.outputContent[lang]);
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.margins = [15, 20, 15, 15];

    var weightRow = optionPanel.add("group");
    weightRow.orientation = "row";
    var showWeightCountCheckbox = weightRow.add("checkbox", undefined, LABELS.showWeightCount[lang]);
    var showWeightCheckbox = weightRow.add("checkbox", undefined, LABELS.showWeightToggle[lang]);
    showWeightCheckbox.value = true;

    var layoutRow = optionPanel.add("group");
    layoutRow.orientation = "row";
    layoutRow.add("statictext", undefined, LABELS.columns[lang]);
    var columnField = layoutRow.add("edittext", undefined, "3");
    columnField.characters = 3;
    var showScoreCheckbox = layoutRow.add("checkbox", undefined, LABELS.showScoreToggle[lang]);

    // --- Insert weight/type filter group panel ---
    var weightGroup = dialog.add("group");
    weightGroup.orientation = "row";
    weightGroup.alignChildren = "top";

    var weightPanel = weightGroup.add("panel", undefined, "ウエイト");
    weightPanel.orientation = "column";
    weightPanel.alignChildren = "left";
    weightPanel.margins = [15, 20, 15, 15];

    var cbVeryThin = weightPanel.add("checkbox", undefined, "超極細・極細");
    var cbLight    = weightPanel.add("checkbox", undefined, "細め");
    var cbRegular  = weightPanel.add("checkbox", undefined, "標準");
    var cbSemiBold = weightPanel.add("checkbox", undefined, "中太");
    var cbBold     = weightPanel.add("checkbox", undefined, "太字・極太");

    var typePanel = weightGroup.add("panel", undefined, "種類");
    typePanel.orientation = "column";
    typePanel.alignChildren = "left";
    typePanel.margins = [15, 20, 15, 15];

    var cbBasic     = typePanel.add("checkbox", undefined, "基本");
    var cbNarrow    = typePanel.add("checkbox", undefined, "狭める系");
    var cbWide      = typePanel.add("checkbox", undefined, "広げる系");
    var cbDecor     = typePanel.add("checkbox", undefined, "装飾・特殊用途");
    var cbSizeProp  = typePanel.add("checkbox", undefined, "サイズ・プロポーション系");

    // --- Insert: updateDisplayPanelState function and event hookup ---
    function updateDisplayPanelState() {
        displayPanel.enabled = true;
        columnField.enabled = true;

        // --- Updated logic for showScoreCheckbox ---
        if (!showWeightCheckbox.value || !displayOptions[0].value) {
            showScoreCheckbox.value = false;
            showScoreCheckbox.enabled = false;
        } else {
            showScoreCheckbox.enabled = true;
        }
    }

    showWeightCheckbox.onClick = updateDisplayPanelState;
    updateDisplayPanelState();

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.margins = [0, 15, 0, 0];
    buttonGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    var result = dialog.show();

    if (result !== 1) return null;

    // Removed "debug" mode; now 5 elements, matching displayOptions
    var displayModeMap = ["family+style", "postscript", "alphabet", "numbers", "custom"];
    var displayMode = displayModeMap[0];
    for (i = 0; i < displayOptions.length; i++) {
        if (displayOptions[i].value) {
            displayMode = displayModeMap[i];
            break;
        }
    }

    return {
        keyword: keywordField.text.toLowerCase().replace(/^\s+|\s+$/g, ""),
        displayMode: displayMode,
        customText: customTextField.text,
        columns: parseInt(columnField.text, 10) || 3,
        useCategory: true,
        showWeight: showWeightCheckbox.value,
        showWeightCount: showWeightCountCheckbox.value,
        showScore: showScoreCheckbox.value,
        weightFilters: {
            veryThin: cbVeryThin.value,
            light: cbLight.value,
            regular: cbRegular.value,
            semiBold: cbSemiBold.value,
            bold: cbBold.value
        },
        typeFilters: {
            basic: cbBasic.value,
            narrow: cbNarrow.value,
            wide: cbWide.value,
            decor: cbDecor.value,
            sizeProp: cbSizeProp.value
        }
    };
}

function confirmShowAllFonts() {
    var confirmDialog = new Window("dialog", LABELS.confirmTitle[lang]);
    confirmDialog.orientation = "column";
    confirmDialog.alignChildren = "left";
    confirmDialog.margins = 20;
    confirmDialog.add("statictext", undefined, LABELS.confirmMessage1[lang]);
    confirmDialog.add("statictext", undefined, LABELS.confirmMessage2[lang]);

    var buttonGroup = confirmDialog.add("group");
    buttonGroup.alignment = "right";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.stop[lang], {
        name: "cancel"
    });
    var okBtn = buttonGroup.add("button", undefined, LABELS.proceed[lang], {
        name: "ok"
    });

    var result = false;
    okBtn.onClick = function() {
        result = true;
        confirmDialog.close();
    };
    cancelBtn.onClick = function() {
        result = false;
        confirmDialog.close();
    };

    confirmDialog.show();
    return result;
}

function mapToFiveRanks(rank) {
    if (rank <= 1) return "veryThin";       // 超極細・極細
    if (rank === 2) return "light";         // 細め
    if (rank >= 3 && rank <= 4) return "regular"; // 標準
    if (rank >= 5 && rank <= 7) return "semiBold"; // 中太
    if (rank >= 8) return "bold";           // 太字・極太
    return "regular";
}

function mapToTypeCategory(style) {
    style = (style || "").toLowerCase();

    // 基本文字カテゴリ
    if (style.indexOf("text") !== -1 || style.indexOf("headline") !== -1) {
        return "basic";
    }

    // 狭める系
    if (style.indexOf("cond") !== -1 || style.indexOf("cn") !== -1 ||
        style.indexOf("compressed") !== -1 || style.indexOf("comp") !== -1) {
        return "narrow";
    }

    // 広げる系
    if (style.indexOf("expanded") !== -1 || style.indexOf("extended") !== -1) {
        return "wide";
    }

    // 装飾・特殊用途
    if (style.indexOf("compact") !== -1 || style.indexOf("display") !== -1) {
        return "decor";
    }

    // サイズ・プロポーション系
    if (style.indexOf("micro") !== -1 || style.indexOf("low") !== -1 || style.indexOf("wide") !== -1) {
        return "sizeProp";
    }

    return null;
}

function collectFonts(input) {
    // Normalization function for quoted search
    function normalize(str) {
        return str.toLowerCase().replace(/[\s\-　]/g, "");
    }

    var grouped = {};
    var i, j;

    var raw = input.keyword;
    var andGroups = [];
    var notKeywords = [];

    if (raw) {
        // -------------------------------
        // 入力の正規化
        // -------------------------------
        raw = raw.replace(/　/g, " "); // 全角スペース→半角
        raw = raw.replace(/\s*,\s*/g, ","); // カンマの前後スペース除去
        raw = raw.replace(/\s+/g, " "); // 連続スペース→1つの半角スペース

        // -------------------------------
        // NOTキーワード（-付き）抽出
        // -------------------------------
        var rawParts = raw.split(" ");
        var keywordParts = [];
        for (i = 0; i < rawParts.length; i++) {
            if (rawParts[i].charAt(0) === "-") {
                notKeywords.push(rawParts[i].substring(1).toLowerCase());
            } else {
                keywordParts.push(rawParts[i]);
            }
        }

        // -------------------------------
        // AND（+区切り）と OR（スペース・カンマ）の構築
        // -------------------------------
        var andParts = keywordParts.join(" ").split("+");
        for (i = 0; i < andParts.length; i++) {
            var group = [];
            var orParts = andParts[i].split(/[\s,]+/);

            for (j = 0; j < orParts.length; j++) {
                var keyword = orParts[j].toLowerCase();
                if (keyword.length === 0) continue;

                var isPrefix = false;
                if (keyword.charAt(0) === "^") {
                    isPrefix = true;
                    keyword = keyword.substring(1);
                }

                // Detect quoted terms and mark them
                var isQuoted = false;
                if (keyword.charAt(0) === '"' && keyword.charAt(keyword.length - 1) === '"') {
                    keyword = keyword.slice(1, -1);
                    isQuoted = true;
                }
                group.push({
                    keyword: keyword,
                    isPrefix: isPrefix,
                    isQuoted: isQuoted
                });
            }
            if (group.length > 0) andGroups.push(group);
        }
    }

    // -------------------------------
    // フォント走査
    // -------------------------------
    for (i = 0; i < textFonts.length; i++) {
        var font = textFonts[i];
        var name = font.name.toLowerCase();
        var family = font.family.toLowerCase();
        var style = (font.style || "").toLowerCase();

        // ✅ NOT 条件：含まれていたら除外
        var skip = false;
        for (j = 0; j < notKeywords.length; j++) {
            var nkey = notKeywords[j];
            if (name.indexOf(nkey) !== -1 || family.indexOf(nkey) !== -1 || style.indexOf(nkey) !== -1) {
                skip = true;
                break;
            }
        }
        if (skip) continue;

        // ✅ AND × OR 条件チェック
        if (andGroups.length > 0) {
            var allMatched = true;

            for (j = 0; j < andGroups.length; j++) {
                var groupMatched = false;
                var group = andGroups[j];

                for (var k = 0; k < group.length; k++) {
                    var kw = group[k];
                    var key = kw.keyword;
                    var isPrefix = kw.isPrefix;

                    if (isPrefix) {
                        if (name.substr(0, key.length) === key ||
                            family.substr(0, key.length) === key ||
                            style.substr(0, key.length) === key) {
                            groupMatched = true;
                            break;
                        }
                    } else {
                        if (kw.isQuoted) {
                            var normalizedKey = normalize(key);
                            if (
                                normalize(name).indexOf(normalizedKey) !== -1 ||
                                normalize(family).indexOf(normalizedKey) !== -1 ||
                                normalize(style).indexOf(normalizedKey) !== -1
                            ) {
                                groupMatched = true;
                                break;
                            }
                        } else {
                            if (name.indexOf(key) !== -1 ||
                                family.indexOf(key) !== -1 ||
                                style.indexOf(key) !== -1) {
                                groupMatched = true;
                                break;
                            }
                        }
                    }
                }

                if (!groupMatched) {
                    allMatched = false;
                    break;
                }
            }

            if (!allMatched) continue;
        }

        // --- Insert weight filter logic here ---
        var rank = getWeightOrderIndex(style, font.name.toLowerCase(), font.family.toLowerCase());
        function classifyRank(rank) {
            if (rank <= 1) return "veryThin";
            if (rank <= 2) return "light";
            if (rank <= 4) return "regular";
            if (rank <= 7) return "semiBold";
            return "bold";
        }
var weightCategory = mapToFiveRanks(rank);
var filters = input.weightFilters;
var anyWeightSelected = filters && (filters.veryThin || filters.light || filters.regular || filters.semiBold || filters.bold);
var weightMatched = !anyWeightSelected || filters[weightCategory];

var typeFilters = input.typeFilters;
var anyTypeSelected = typeFilters && (typeFilters.basic || typeFilters.narrow || typeFilters.wide || typeFilters.decor || typeFilters.sizeProp);
var typeCategory = mapToTypeCategory(style);
var typeMatched = !anyTypeSelected || (typeCategory && typeFilters[typeCategory]);

// --- OR条件に変更 ---
// 両方未選択 → 全件許可
// どちらかマッチすれば残す
if (!weightMatched && !typeMatched) {
    continue;
}
        // --- End weight filter logic ---

        // --- Insert type filter logic here ---
        var typeFilters = input.typeFilters;
        var anyTypeSelected = typeFilters && (typeFilters.basic || typeFilters.narrow || typeFilters.wide || typeFilters.decor || typeFilters.sizeProp);
        if (anyTypeSelected) {
            var typeMatched = false;

            // classify type based on flags
            if (typeFilters.basic && (style.indexOf("text") !== -1 || style.indexOf("headline") !== -1)) {
                typeMatched = true;
            }
            if (typeFilters.narrow && (style.indexOf("cond") !== -1 || style.indexOf("cn") !== -1 || style.indexOf("compressed") !== -1 || style.indexOf("comp") !== -1)) {
                typeMatched = true;
            }
            if (typeFilters.wide && (style.indexOf("expanded") !== -1 || style.indexOf("extended") !== -1)) {
                typeMatched = true;
            }
            if (typeFilters.decor && (style.indexOf("compact") !== -1 || style.indexOf("display") !== -1)) {
                typeMatched = true;
            }
            if (typeFilters.sizeProp && (style.indexOf("micro") !== -1 || style.indexOf("low") !== -1 || style.indexOf("wide") !== -1)) {
                typeMatched = true;
            }

            if (!typeMatched) {
                continue;
            }
        }
        // --- End type filter logic ---

        // ✅ グループに追加
        var groupKey = input.useCategory ? font.family : "Uncategorized";
        if (!grouped[groupKey]) grouped[groupKey] = [];

        var exists = false;
        for (j = 0; j < grouped[groupKey].length; j++) {
            if (grouped[groupKey][j].name === font.name) {
                exists = true;
                break;
            }
        }

        if (!exists) grouped[groupKey].push(font);
    }

    return grouped;
}

// --------------------------------
// 各グループをウエイト＋スタイル順にソート（ES3）
// --------------------------------
function sortFontGroups(groups) {
    for (var label in groups) {
        if (!groups.hasOwnProperty(label)) continue;
        groups[label].sort(function(a, b) {
            return getFullSortRank(a) - getFullSortRank(b);
        });
    }
}
// --------------------------------
// アートボード上にフォントサンプルを描画（ES3）
// --------------------------------
function drawFontSamples(doc, groupedFonts, input) {
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var bounds = artboard.artboardRect;
    var startX = bounds[0] + 20;
    var startY = bounds[1] - 20;
    var columnSpacing = 220;
    var rowSpacing = 300;
    var fontSize = 10;

    app.executeMenuCommand("deselectall");

    // グループ名をソート取得（ES3）
    var groupLabels = [];
    for (var label in groupedFonts) {
        if (!groupedFonts.hasOwnProperty(label)) continue;
        if (groupedFonts[label].length > 0) {
            groupLabels.push(label);
        }
    }

    // ソート（ES3対応）
    groupLabels.sort();

    var col = 0;
    var row = 0;
    for (var i = 0; i < groupLabels.length; i++) {
        var label = groupLabels[i];
        var fonts = groupedFonts[label];

        var x = startX + col * columnSpacing;
        var y = startY - row * rowSpacing;

        if (++col >= input.columns) {
            col = 0;
            row++;
        }

        if (input.showWeight) {
            if (input.useCategory) {
                var labelFrame = doc.textFrames.add();
                labelFrame.contents = "[" + label + "]" + (input.showWeightCount ? " (" + fonts.length + ")" : "");
                labelFrame.left = x;
                labelFrame.top = y;
                labelFrame.selected = true;
                labelFrame.textRange.characterAttributes.size = fontSize;
                y -= labelFrame.height + fontSize * 0.5;
            }

            for (var j = 0; j < fonts.length; j++) {
                var font = fonts[j];
                try {
                    var tf = doc.textFrames.add();
                    tf.contents = getDisplayText(font, input.displayMode, input.customText, input.showScore);
                    tf.textRange.characterAttributes.size = fontSize;
                    var safeFont = textFonts.getByName(font.name);
                    tf.textRange.characterAttributes.textFont = safeFont;
                    tf.left = x;
                    tf.top = y;
                    tf.selected = true;
                    y -= tf.height;
                } catch (e) {
                    $.writeln("描画失敗：" + font.name + " → " + e);
                }
            }
        } else {
            x = startX;
            y = startY - i * 16;

            var sampleFont = fonts[0]; // use first font as representative
            var labelFrame = doc.textFrames.add();
            labelFrame.contents = getDisplayText(sampleFont, input.displayMode, input.customText, input.showScore);
            labelFrame.left = x;
            labelFrame.top = y;
            labelFrame.selected = true;
            labelFrame.textRange.characterAttributes.size = fontSize;

            try {
                var safeFont = textFonts.getByName(sampleFont.name);
                labelFrame.textRange.characterAttributes.textFont = safeFont;
                labelFrame.textRange.characterAttributes.size = fontSize;
            } catch (e) {
                $.writeln("フォント適用失敗：" + sampleFont.name + " → " + e);
            }
        }
    }
}

function getSafeFont(name) {
    try {
        return textFonts.getByName(name);
    } catch (e) {
        return null;
    }
}

// --------------------------------
// 表示するテキスト内容を決定（ES3）
// --------------------------------
function getDisplayText(font, mode, customText, showScore) {
    var rank = getFullSortRank(font); // スコア取得
    if (mode === "postscript") {
        return font.name;
    } else if (mode === "alphabet") {
        return "The quick brown fox jumps over the lazy dog.";
    } else if (mode === "numbers") {
        return "1234567890";
    } else if (mode === "family+style") {
        var text = font.family + (font.style ? " " + font.style : "");
        if (showScore) text += " (" + rank + ")";
        return text;
    } else if (mode === "custom") {
        return customText || "";
    } else {
        var text = font.family + (font.style ? " " + font.style : "");
        if (showScore) text += " (" + rank + ")";
        return text;
    }
}
// --------------------------------
// ウエイト・スタイル評価値を取得（加点ロジック付き）
// --------------------------------

function getFullSortRank(font) {
    var style = (font.style || "").toLowerCase();
    var postscriptName = (font.name || "").toLowerCase();
    var fontFamily = (font.family || "").toLowerCase();

    // ✅ 特例：PostScript名が「FuturaPT-Heavy」の場合 → 1015 固定（加点処理なし）
    if (postscriptName === "futurapt-heavy") {
        return 1015;
    }

    var baseRank = getWeightOrderIndex(style, postscriptName, fontFamily);
    var offset = 0;

    var words = style.split(/\s+/);

    // 装飾フラグ初期化
    var flags = {
        hasText: false,
        hasHeadline: false,
        hasCondensed: false,
        hasCn: false,
        hasExpanded: false,
        hasExtended: false,
        hasUltraCondensed: false,
        hasExtraCondensed: false,
        hasSemiCondensed: false,
        hasCompressed: false,
        hasExtraCompressed: false,
        hasCompact: false,
        hasDisplay: false,
        hasMicro: false,
        hasLow: false,
        hasWide: false
    };

    // 装飾キーワードに応じたフラグ設定
    for (var i = 0; i < words.length; i++) {
        var w = words[i];
        if (w === "text") flags.hasText = true;
        if (w === "headline") flags.hasHeadline = true;
        if (w === "cond" || w === "condensed") flags.hasCondensed = true;
        if (w === "cn") flags.hasCn = true;
        if (w === "expanded") flags.hasExpanded = true;
        if (w === "extended") flags.hasExtended = true;
        if (w === "semiextended" || (w === "semi" && words[i + 1] === "extended")) flags.hasExtended = true;
        // --- Begin: Add logic to detect "semiexpanded" and "semi expanded" ---
        if (w === "semiexpanded" || (w === "semi" && words[i + 1] === "expanded")) flags.hasExpanded = true;
        // --- End: Add logic for semiexpanded ---
        if (w === "ultracondensed" || (w === "ultra" && words[i + 1] === "condensed")) flags.hasUltraCondensed = true;
        if (w === "extracondensed" || (w === "extra" && words[i + 1] === "condensed")) flags.hasExtraCondensed = true;
        if (w === "semicondensed" || (w === "semi" && words[i + 1] === "condensed")) flags.hasSemiCondensed = true;
        if (w === "compressed" || w === "comp") flags.hasCompressed = true;
        if (w === "extra" && words[i + 1] === "compressed") flags.hasExtraCompressed = true;
        if (w === "compact") flags.hasCompact = true;
        if (w === "display") flags.hasDisplay = true;
        if (w === "micro") flags.hasMicro = true;
        if (w === "low") flags.hasLow = true;
        if (w === "wide") flags.hasWide = true;
    }

    // Italic 判定（全体 style に対して）
    var isItalic = /italic|oblique|slanted|inclined|kursiv|\bit\b/.test(style);

    // 加点処理（100刻み + 特例あり）
    if (flags.hasDisplay) offset += 100;
    if (flags.hasCompressed) offset += 200;
    if (flags.hasCompact) offset += 300;
    if (flags.hasExpanded) offset += 400;
    if (flags.hasExtended) offset += 500;
    if (flags.hasUltraCondensed) offset += 600;
    if (flags.hasExtraCondensed) offset += 700;
    if (flags.hasSemiCondensed) offset += 850;

    // Condensed系代表加点（複数条件一致でも一度のみ）
    if (
        flags.hasCondensed ||
        flags.hasCn ||
        flags.hasWide ||
        flags.hasSemiCondensed ||
        flags.hasExtraCompressed
    ) {
        offset += 900;
    }

    if (flags.hasHeadline) offset += 1000;
    if (flags.hasText) offset += 1100;
    if (flags.hasLow) offset += 1200;
    if (flags.hasMicro) offset += 1250;
    if (flags.hasWide) offset += 1275;
    if (flags.hasExtraCompressed) offset += 150; // 特別加点
    if (isItalic) offset += 1300;

    return baseRank + offset;
}

// --------------------------------
// スタイル文字列に対する基本ウエイトスコア取得
// --------------------------------

function getWeightOrderIndex(rawStyle, postscriptName, fontFamily) {
    rawStyle = rawStyle || "";
    var style = rawStyle.toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    var words = style.split(/\s+/);
    var i, j, k;

    var weightGroups = [
        ["hairline"], // +0
        ["ultra thin", "ultrathin", "ut"], // +1
        ["thin", "th"], // +2
        ["default"], // +3
        ["ultralight", "ultra light", "ultlt", "ul"], // +4
        ["extralight", "extra light", "el", "xlight", "xl"], // +5
        ["lightsemi"], // +6
        ["light", "lt", "lite", "l"], // +7
        ["lb"], // +8
        ["book", "bk"], // +9
        ["n", "normal"], // +10
        ["middle"], // +11
        ["regular", "roman", "normal", "レギュラー", "r"], // +12
        ["rb"], // +13
        ["medium", "md", "ミディアム", "m"], // +14
        ["semibold", "semi bold", "sb"], // +15
        ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"], // +16
        ["bold", "bd", "ボールド", "b"], // +17
        ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"], // +18
        ["heavy", "h"], // +19
        ["black"], // +20
        ["xblack", "extra black", "extrablack", "xb"], // +21
        ["ultra", "u", "ub", "ultra black", "ultrablack"] // +22
    ];

    // Regular 扱いする単独語句
    var regularSingles = [
        "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
        "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
        "headline", "text", "low", "micro", "extra compressed",
        "semi expanded", "semiexpanded"
    ];

    function getRegularIndex(groups) {
        for (var i = 0; i < groups.length; i++) {
            for (var j = 0; j < groups[i].length; j++) {
                if (groups[i][j] === "regular") return i;
            }
        }
        return 12; // fallback
    }
    var regularIndex = getRegularIndex(weightGroups);
    var applyFrutigerCorrection = (/frutiger/i.test(fontFamily) && /ultralight/.test(style));

    // ✅ W0〜W9
    var wMatch = style.match(/^w(\d)$/);
    if (wMatch !== null) return parseInt(wMatch[1], 10);

    // ✅ W000〜W999
    var w3Match = style.match(/^w(\d{3})$/);
    if (w3Match !== null) return parseInt(w3Match[1], 10);

    // ✅ 先頭数値（例：25 Ultra Light）
    var match = style.match(/^(\d{1,3})(?=\D|$)/);
    if (match) return parseInt(match[1], 10);

    // ✅ 特例：HelveticaNeue, Tazugane + Ultra Light → 999
if (
  (
    /helveticaneue/i.test(postscriptName) ||
    /tazugane/i.test(postscriptName) ||
    /universnextpro/i.test(postscriptName)
  ) &&
  /ultralight|ultra light|ultlt/i.test(style)
) {
  return 999;
}

    // ✅ 単独語が italic / oblique / wide → Regular 扱い
    if (words.length === 1 && /^(italic|oblique|it|wide)$/.test(words[0])) {
        return 1000 + regularIndex;
    }

    // ✅ 装飾語だけなら Regular 扱い
    if (words.length === 1) {
        for (i = 0; i < regularSingles.length; i++) {
            if (words[0] === regularSingles[i]) return 1000 + regularIndex;
        }
    }

    // ✅ 完全一致
    for (i = 0; i < weightGroups.length; i++) {
        for (j = 0; j < weightGroups[i].length; j++) {
            if (style === weightGroups[i][j]) {
                var base = 1000 + i;
                if (applyFrutigerCorrection && i === 4) base -= 5;
                return base;
            }
        }
    }

    // ✅ 複合語一致（長い語優先）＋単語完全一致
    var allTerms = [];
    for (i = 0; i < weightGroups.length; i++) {
        for (j = 0; j < weightGroups[i].length; j++) {
            allTerms.push({
                term: weightGroups[i][j],
                index: i
            });
        }
    }
    // 複合語一致：style 全体に対して長い語から順に indexOf で検出（単語境界を含めて）
    allTerms.sort(function(a, b) {
        return b.term.length - a.term.length;
    });
    for (i = 0; i < allTerms.length; i++) {
        var term = allTerms[i].term;
        var pattern = new RegExp("\\b" + term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "\\b");
        if (pattern.test(style)) {
            var base = 1000 + allTerms[i].index;
            if (applyFrutigerCorrection && allTerms[i].index === 4) base -= 5;
            return base;
        }
    }

    // ✅ fallback：Regular 扱い
    var fallback = 1000 + regularIndex;
    if (applyFrutigerCorrection) fallback -= 5;
    return fallback;
}

function main() {
    try {
        if (app.documents.length === 0) {
            alert(LABELS.errorNoDocument[lang]);
            return;
        }

        var doc = app.activeDocument;
        var userInput = showFontListDialog();
        if (!userInput) return;

        if (!userInput.keyword && !confirmShowAllFonts()) return;

        var fontGroups = collectFonts(userInput);
        sortFontGroups(fontGroups);
        // --- Inserted logic: show alert if no matching fonts found ---
        var totalFonts = 0;
        for (var key in fontGroups) {
            if (fontGroups.hasOwnProperty(key)) {
                totalFonts += fontGroups[key].length;
            }
        }
        if (totalFonts === 0) {
            alert("条件に該当するフォントが見つかりませんでした。");
            return;
        }
        // -----------------------------------------------------------
        drawFontSamples(doc, fontGroups, userInput);
    } catch (e) {
        alert(LABELS.errorOccurred[lang] + e.message);
    }
}

main(); // スクリプト実行開始