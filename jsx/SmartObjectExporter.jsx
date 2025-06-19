#target illustrator

/*
 * スクリプト名：SmartObjectExporter.jsx
 *
 * 概要：
 * Adobe Illustrator の選択オブジェクトを一時アートボードとプレビュー用レイヤーに複製し、
 * 背景・余白・罫線・書き出し倍率・ファイル名構成などを指定して PNG ファイルとして書き出します。
 * ダイアログではリアルタイムプレビューやプリセット保存が可能です。
 *
 * 【主な機能】
 * - 選択オブジェクトを一時アートボードに複製して PNG 書き出し
 * - 背景設定（透過・白・黒・任意カラー・透明グリッド）
 * - 書き出し倍率または幅指定（1x〜4x／カスタム倍率／px幅）
 * - 余白設定（なし／左右／上下／四辺）
 * - 罫線の追加（色・太さ指定、黒／白／任意カラー）
 * - ファイル名構成（区切り記号・接尾辞）＋リアルタイムプレビュー
 * - 書き出し先の選択（デスクトップ／元ファイルのフォルダ）
 * - 書き出し後に Finder でフォルダ表示（Mac のみ）
 * - 設定のプリセット保存（.txt ファイル形式）
 * - プレビュー用レイヤー「__preview」は常に削除（キャンセル時含む）
 * - レイヤー順を維持したままプレビュー用レイヤーへ複製
 * - 元の非表示レイヤーを自動で再表示
 *
 * 更新履歴：
 * - 0.5.0  : 初版作成
 * - 0.5.13 : 選択オブジェクトの重ね順を維持したまま複製（PLACEATEND 使用）
 * - 0.5.14 : 余白オプション（左右／上下／四辺）を追加
 * - 更新日：2025-05-27
 */
 

// -------------------------------
// Helper: Restore layer visibility
// -------------------------------
function restoreLayerVisibility(layersArray) {
    for (var i = 0; i < layersArray.length; i++) {
        try {
            layersArray[i].visible = true;
        } catch (e) {
            alert("Error restoring layer visibility: " + e.message);
        }
    }
}
// -------------------------------
// Helper: Hide all layers except one
// -------------------------------
function hideOtherLayers(doc, exceptLayer) {
    var hidden = [];
    for (var i = 0; i < doc.layers.length; i++) {
        var lyr = doc.layers[i];
        if (lyr !== exceptLayer && lyr.visible) {
            try {
                lyr.visible = false;
                hidden.push(lyr);
            } catch (e) {
                alert("Error hiding layer: " + e.message);
            }
        }
    }
    return hidden;
}
// -------------------------------
// Helper: Duplicate selection to a target layer (重ね順維持: PLACEATEND)
// -------------------------------
function duplicateSelectionToLayer(selection, targetLayer) {
    var duplicatedItems = [];
    for (var i = 0; i < selection.length; i++) {
        try {
            var duplicated = selection[i].duplicate(targetLayer, ElementPlacement.PLACEATEND);
            duplicatedItems.push(duplicated);
        } catch (e) {
            alert("Error duplicating selection to preview layer: " + e.message);
        }
    }
    return duplicatedItems;
}
// -------------------------------
// Helper: Set suffix from value and update filename preview
// (Always assign value to suffixInput.text even if value === "")
// -------------------------------
function setSuffixFrom(value) {
    if (typeof rbNoSuffix !== "undefined" && rbNoSuffix !== null) {
        rbNoSuffix.value = false;
    }
    if (typeof rbCustomSuffix !== "undefined" && rbCustomSuffix !== null) {
        rbCustomSuffix.value = true;
    }
    if (typeof suffixInput !== "undefined" && suffixInput !== null) {
        suffixInput.enabled = true;
        suffixInput.text = value;
    }
    updateFilenamePreview();
}

var PREVIEW_LAYER_NAME = "__preview";

var docName = "";

// -------------------------------
// 言語判定関数 Language detection
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// 日英ラベル定義 Define UI labels
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    background: {
        ja: "背景色",
        en: "Background"
    },
    black: {
        ja: "黒",
        en: "Black"
    },
    blackLine: {
        ja: "黒",
        en: "Black"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    custom: {
        ja: "カスタム",
        en: "Custom"
    },
    customScaleLabel: {
        ja: "倍率：",
        en: "Scale: "
    },
    desktop: {
        ja: "デスクトップ",
        en: "Desktop"
    },
    dialogTitle: {
        ja: "選択オブジェクトを書き出し",
        en: "Export Selected Objects"
    },
    exportFailed: {
        ja: "書き出しに失敗しました：",
        en: "Export failed: "
    },
    exportPreset: {
        ja: "プリセットを保存",
        en: "Save Preset"
    },
    fileName: {
        ja: "書き出しファイル名",
        en: "Export Filename"
    },
    fileNameReference: {
        ja: "ファイル名",
        en: "Filename"
    },
    hexColor: {
        ja: "カラー指定",
        en: "Color Code"
    },
    invalidSize: {
        ja: "選択範囲のサイズが無効です。",
        en: "Invalid selection size."
    },
    location: {
        ja: "書き出し先",
        en: "Export Location"
    },
    margin: {
        ja: "余白",
        en: "Margin"
    },
    marginHorizontal: {
        ja: "左右",
        en: "Horizontal"
    },
    marginVertical: {
        ja: "上下",
        en: "Vertical"
    },
    marginAll: {
        ja: "四辺",
        en: "All Sides"
    },
    noMargin: {
        ja: "つけない",
        en: "None"
    },
    none: {
        ja: "なし",
        en: "None"
    },
    noRule: {
        ja: "つけない",
        en: "None"
    },
    noSelection: {
        ja: "ドキュメントが開かれていないか、オブジェクトが選択されていません。",
        en: "No document open or no object selected."
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    presetLabel: {
        ja: "プリセット：",
        en: "Preset: "
    },
    presetPrompt: {
        ja: "プリセット名を入力してください",
        en: "Enter preset name"
    },
    presetSavedMessage: {
        ja: "プリセットを保存しました：",
        en: "Preset saved: "
    },
    rule: {
        ja: "罫線",
        en: "Border"
    },
    ruleColor: {
        ja: "罫線カラー",
        en: "Border Color"
    },
    sameFolder: {
        ja: "ファイルと同じ場所",
        en: "Same as File"
    },
    showFolder: {
        ja: "書き出し後、フォルダーを表示",
        en: "Show Folder After Export"
    },
    size: {
        ja: "書き出しサイズ（px）",
        en: "Export Size (px)"
    },
    suffix: {
        ja: "接尾辞：",
        en: "Suffix: "
    },
    suffixTextLabel: {
        ja: "接尾辞",
        en: "Suffix"
    },
    symbol: {
        ja: "区切り文字",
        en: "delimiter"
    },
    transparent: {
        ja: "透過",
        en: "Transparent"
    },
    transparentGrid: {
        ja: "透明グリッド",
        en: "Transparency Grid"
    },
    white: {
        ja: "白",
        en: "White"
    },
    whiteLine: {
        ja: "白",
        en: "White"
    },
    widthInput: {
        ja: "横幅：",
        en: "Width: "
    },
    withRule: {
        ja: "つける",
        en: "Add"
    },
    // デフォルトプリセット名（ハードコードから移動）
    defaultPresetName: {
        ja: "マイプリセット",
        en: "MyPreset"
    }
};
// -------------------------------
// Helper: Restore temporary hidden items
// -------------------------------
function restoreTemporaryHiddenItems(hiddenItems) {
    for (var i = 0; i < hiddenItems.length; i++) {
        try {
            if (hiddenItems[i] && hiddenItems[i].isValid) {
                hiddenItems[i].hidden = false;
            }
        } catch (e) {
            alert("Error restoring temporary hidden item: " + e.message);
        }
    }
}

// 構造化されたプリセット定義（仕様準拠: background, margin, stroke, location, symbol, suffix, size）
var presets = [
    {
        label: (lang === 'ja') ? "透過・余白なし・倍率200%" : "Transparent / No Margin / 200%",
        background: "transparent",
        margin: "none",
        stroke: "none",
        location: "desktop",
        symbol: "",
        suffix: "200",
        size: "scale:200"
    },
    {
        label: (lang === 'ja') ? "白背景・左右3mm・倍率300%" : "White BG / Horizontal 3mm / 300%",
        background: "white",
        margin: "horizontal:3",
        stroke: "none",
        location: "desktop",
        symbol: "_",
        suffix: "300",
        size: "scale:300"
    },
    {
        label: (lang === 'ja') ? "白背景・上下5mm・幅指定1000px" : "White BG / Vertical 5mm / Width 1000px",
        background: "white",
        margin: "vertical:5",
        stroke: "1,black",
        location: "desktop",
        symbol: "_",
        suffix: "1000",
        size: "width:1000"
    },
    {
        label: (lang === 'ja') ? "黒背景・四辺10mm・倍率200%" : "Black BG / All 10mm / 200%",
        background: "black",
        margin: "all:10",
        stroke: "none",
        location: "desktop",
        symbol: "-",
        suffix: "200",
        size: "scale:200"
    }
];

// ES3安全: プリセットラベル配列を手動で作成
var presetLabels = [];
for (var i = 0; i < presets.length; i++) {
    presetLabels.push(presets[i].label);
}

function showExportOptionsDialog(selectionBounds, temporaryHiddenItems) {
    var previewHiddenItems = [];
    // --- Helper: Get current background value ---
    function getCurrentBackgroundValue() {
        if (rbTransparentGrid.value) return "transparentGrid";
        if (rbWhite.value) return "white";
        if (rbBlack.value) return "black";
        if (rbHex.value) {
            var hex = hexInput.text;
            return (hex.charAt(0) === "#" ? hex : "#" + hex);
        }
        return "transparent";
    }

    // --- Helper: Get current margin value ---
    function getCurrentMarginValue() {
        if (rbNoMargin.value) return "none";
        if (rbMarginH.value) return "horizontal:" + marginInput.text;
        if (rbMarginV.value) return "vertical:" + marginInput.text;
        if (rbMarginAll.value) return "all:" + marginInput.text;
        return "none";
    }

    // --- Helper: Get current stroke value ---
    function getCurrentStrokeValue() {
        return rbNoRule.value ? "none" : ruleInput.text + "," + (rbBlackLine.value ? "black" : "white");
    }

    // --- Helper: Get current size value ---
    function getCurrentSizeValue() {
        if (rbScale100.value) return "scale:100";
        if (rbScale200.value) return "scale:200";
        if (rbScale300.value) return "scale:300";
        if (rbScale400.value) return "scale:400";

        if (rbCustomScale.value) return "scale:" + customScaleInput.text;
        if (rbWidthInput.value) return "width:" + widthInput.text;
        return "scale:100";
    }
    // --- Helper: Clear all scale radio selections ---
    function clearAllScaleRadioSelections() {
        var radios = [rbScale100, rbScale200, rbScale300, rbScale400, rbCustomScale, rbWidthInput];
        for (var i = 0; i < radios.length; i++) {
            if (typeof radios[i] !== "undefined" && radios[i] !== null) {
                radios[i].value = false;
            }
        }
    }

    // --- Preview apply function ---
    function applyPreviewFromUI() {
        if (typeof previewApply === "function") {
            previewApply();
        }
    }

    function toPx(val) {
        return Math.ceil(val);
    }

    var unitInfo = getRulerUnitInfo();
    var unitLabel = unitInfo.label;
    var unitFactor = unitInfo.factor;
    var defaultMargin = (unitLabel === "mm") ? 3 : 10;

    var width = (selectionBounds[2] - selectionBounds[0]);
    var height = (selectionBounds[1] - selectionBounds[3]);

    // Labels for scale radio buttons: initially exclude pixel info
    var label100 = "1x";
    var label200 = "2x";
    var label300 = "3x";
    var label400 = "4x";

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);

    // (Dialog location logic removed)
    dialog.alignChildren = "left";

    // 上部にプリセット選択を追加
    var presetGroup = dialog.add("group");
    presetGroup.margins = [16, 10, 10, 15];
    presetGroup.orientation = "row";
    presetGroup.alignment = "left";
    presetGroup.add("statictext", undefined, LABELS.presetLabel[lang]);

    var presetDropdown = presetGroup.add("dropdownlist", undefined, [
        LABELS.custom[lang]
    ].concat(presetLabels));
    presetDropdown.selection = 0;

    // --- プリセット書き出しボタンを追加 ---
    var btnExportPreset = presetGroup.add("button", undefined, LABELS.exportPreset[lang]);

    // (btnExportPreset.onClick handler will be inserted later)
    // --- Inserted btnExportPreset.onClick handler after UI variable definitions ---
    btnExportPreset.onClick = function() {
        try {
            var presetName = prompt(LABELS.presetPrompt[lang], LABELS.defaultPresetName[lang]);
            if (!presetName) return;

            var now = new Date();
            var yyyyMMdd = now.getFullYear().toString() +
                ("0" + (now.getMonth() + 1)).slice(-2) +
                ("0" + now.getDate()).slice(-2);
            var baseName = "export-setting-" + yyyyMMdd;
            var counter = 1;
            var exportFile = new File(Folder.desktop + "/" + baseName + ".txt");
            while (exportFile.exists) {
                counter++;
                exportFile = new File(Folder.desktop + "/" + baseName + "_" + counter + ".txt");
            }

            var presetText = "";
            presetText += ",\n";
            presetText += "{label: (lang === 'ja') ? \"" + presetName + "\" : \"\",\n";
            // --- Use helper functions for background, margin, stroke ---
            presetText += "background=\"" + getCurrentBackgroundValue() + "\",\n";
            presetText += "margin=\"" + getCurrentMarginValue() + "\",\n";
            presetText += "stroke=\"" + getCurrentStrokeValue() + "\",\n";
            presetText += "location=\"" + (rbDesktop.value ? "desktop" : "sameFolder") + "\",\n";

            var symbol = "";
            if (rbDash.value) {
                symbol = "-";
            } else if (rbUnderscore.value) {
                symbol = "_";
            }
            presetText += "symbol=\"" + symbol + "\",\n";
            presetText += "suffix=\"" + suffixInput.text + "\"\n";
            // --- Use helper for size ---
            presetText += "size=\"" + getCurrentSizeValue() + "\"\n";
            presetText += "}";

            exportFile.encoding = "UTF-8";
            exportFile.open("w");
            exportFile.write(presetText);
            exportFile.close();
            alert(LABELS.presetSavedMessage[lang] + exportFile.name);
        } catch (e) {
            alert("プリセット保存時にエラーが発生しました: " + e.message);
        }
    };

    // 2カラムのメイングループ
    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";

    var colLeft = mainGroup.add("group");
    colLeft.orientation = "column";
    colLeft.spacing = 16;
    colLeft.alignChildren = "left";
    colLeft.alignment = "top";

    var colRight = mainGroup.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = "left";
    colRight.alignment = "top";

    // 左カラム: 背景色（ラジオボタン: 透過・白・黒を横並び, カラー指定は下）
    var bgPanel = colLeft.add("panel", undefined, LABELS.background[lang]);
    bgPanel.margins = [16, 20, 10, 10];
    bgPanel.orientation = "column";
    bgPanel.alignChildren = "left";


    // 背景色ラジオボタンを横並びにする（順序: 透過, 黒, 白）
    var bgGroup = bgPanel.add("group");
    bgGroup.orientation = "row";
    var rbTransparent = bgGroup.add("radiobutton", undefined, LABELS.transparent[lang]);
    var rbBlack = bgGroup.add("radiobutton", undefined, LABELS.black[lang]);
    var rbWhite = bgGroup.add("radiobutton", undefined, LABELS.white[lang]);
    // 透明グリッド（縦に追加）＋サイズ入力欄
    var transparentGridGroup = bgPanel.add("group");
    transparentGridGroup.orientation = "row";
    var rbTransparentGrid = transparentGridGroup.add("radiobutton", undefined, LABELS.transparentGrid[lang]);
    var transparentGridInput = transparentGridGroup.add("edittext", undefined, "100");
    transparentGridInput.characters = 4;
    transparentGridInput.enabled = false;
    var transparentGridLabel = transparentGridGroup.add("statictext", undefined, "%");
    rbTransparent.value = true;
    // --- 透明グリッドのパーセンテージ値変更時に即プレビュー反映（プレビューを即時更新） ---
    transparentGridInput.onChange = function() {
        applyPreviewFromUI();
    };

    // 背景色: Hex カラー指定（ラジオ＋テキスト）
    var hexGroup = bgPanel.add("group");
    hexGroup.orientation = "row";
    hexGroup.alignChildren = "left";
    var rbHex = hexGroup.add("radiobutton", undefined, LABELS.hexColor[lang]);
    var hexInput = hexGroup.add("edittext", undefined, "#ffcc00");
    hexInput.characters = 12;
    hexInput.enabled = false;

    // --- Insert exclusive background selection logic here ---
    // 背景色選択ラジオ排他制御（選択以外をオフにして入力欄有効化も制御）
    function clearBackgroundSelection(except) {
        rbTransparent.value = (except === rbTransparent);
        rbWhite.value = (except === rbWhite);
        rbBlack.value = (except === rbBlack);
        rbHex.value = (except === rbHex);
        rbTransparentGrid.value = (except === rbTransparentGrid);

        hexInput.enabled = (except === rbHex);
        transparentGridInput.enabled = (except === rbTransparentGrid);

        applyPreviewFromUI();
    }

    // Refactored: Use a loop to assign onClick handlers for background radio buttons
    var bgRadios = [rbTransparent, rbWhite, rbBlack, rbHex, rbTransparentGrid];
    for (var i = 0; i < bgRadios.length; i++) {
        (function(btn) {
            btn.onClick = function() {
                clearBackgroundSelection(btn);
            };
        })(bgRadios[i]);
    }
    hexInput.onChange = applyPreviewFromUI;

    // 左カラム: 余白
    var marginPanel = colLeft.add("panel", undefined, LABELS.margin[lang]);
    marginPanel.margins = [16, 20, 10, 10];
    marginPanel.orientation = "column";
    marginPanel.alignChildren = "left";

    var rbNoMargin = marginPanel.add("radiobutton", undefined, LABELS.noMargin[lang]);
    // 横並びで余白ラジオボタン3つ
    var marginRadioGroup = marginPanel.add("group");
    marginRadioGroup.orientation = "row";
    var rbMarginH = marginRadioGroup.add("radiobutton", undefined, LABELS.marginHorizontal[lang]);
    var rbMarginV = marginRadioGroup.add("radiobutton", undefined, LABELS.marginVertical[lang]);
    var rbMarginAll = marginRadioGroup.add("radiobutton", undefined, LABELS.marginAll[lang]);
    rbNoMargin.value = true;

    var marginGroup = marginPanel.add("group");
    marginGroup.orientation = "row";
    var marginInput = marginGroup.add("edittext", undefined, String(defaultMargin));
    marginInput.characters = 4;
    var marginLabel = marginGroup.add("statictext", undefined, unitLabel);
    marginInput.enabled = false;

    rbNoMargin.onClick = function () {
        // 排他処理
        rbMarginH.value = false;
        rbMarginV.value = false;
        rbMarginAll.value = false;
        rbNoMargin.value = true;
        marginInput.enabled = false;
        applyPreviewFromUI();
    };
    rbMarginH.onClick = function () {
        rbNoMargin.value = false;
        rbMarginV.value = false;
        rbMarginAll.value = false;
        rbMarginH.value = true;
        marginInput.enabled = true;
        applyPreviewFromUI();
    };
    rbMarginV.onClick = function () {
        rbNoMargin.value = false;
        rbMarginH.value = false;
        rbMarginAll.value = false;
        rbMarginV.value = true;
        marginInput.enabled = true;
        applyPreviewFromUI();
    };
    rbMarginAll.onClick = function () {
        rbNoMargin.value = false;
        rbMarginH.value = false;
        rbMarginV.value = false;
        rbMarginAll.value = true;
        marginInput.enabled = true;
        applyPreviewFromUI();
    };
    marginInput.onChange = applyPreviewFromUI;

    // 左カラム: 罫線
    var rulePanel = colLeft.add("panel", undefined, LABELS.rule[lang]);
    rulePanel.margins = [16, 20, 10, 10];
    rulePanel.orientation = "column";
    rulePanel.alignChildren = "left";

    var rbNoRule = rulePanel.add("radiobutton", undefined, LABELS.noRule[lang]);
    var ruleGroup = rulePanel.add("group");
    ruleGroup.orientation = "row";
    var rbWithRule = ruleGroup.add("radiobutton", undefined, LABELS.withRule[lang]);
    // --- Add exclusive event handler for rbWithRule ---
    rbWithRule.onClick = function() {
        rbNoRule.value = false;
        ruleInput.enabled = true;
        ruleColorLabelGroup.enabled = true;
        ruleColorRadioGroup.enabled = true;
        ruleHexGroup.enabled = true;
    };
    var ruleDefault = (unitLabel === "mm") ? "0.1" : "1";
    var ruleInput = ruleGroup.add("edittext", undefined, ruleDefault);
    ruleInput.characters = 4;
    ruleInput.enabled = false;
    var ruleLabel = ruleGroup.add("statictext", undefined, unitLabel);

    rbNoRule.value = true;

    // 罫線カラーグループを罫線パネル内の最後に追加（ラベル・ラジオ・Hex入力を分割）
    var ruleColorLabelGroup = rulePanel.add("group");
    ruleColorLabelGroup.orientation = "row";
    ruleColorLabelGroup.enabled = false;
    ruleColorLabelGroup.add("statictext", undefined, LABELS.ruleColor[lang]);

    var ruleColorRadioGroup = rulePanel.add("group");
    ruleColorRadioGroup.orientation = "row";
    ruleColorRadioGroup.enabled = false;
    var rbBlackLine = ruleColorRadioGroup.add("radiobutton", undefined, LABELS.blackLine[lang]);
    var rbWhiteLine = ruleColorRadioGroup.add("radiobutton", undefined, LABELS.whiteLine[lang]);

    var ruleHexGroup = rulePanel.add("group");
    ruleHexGroup.orientation = "row";
    ruleHexGroup.enabled = false;
    var rbRuleHex = ruleHexGroup.add("radiobutton", undefined, LABELS.hexColor[lang]);
    var ruleHexInput = ruleHexGroup.add("edittext", undefined, "#333333");
    ruleHexInput.characters = 12;
    ruleHexInput.enabled = false;

    rbBlackLine.value = true;

    rbNoRule.onClick = function() {
        rbWithRule.value = false;
        ruleInput.enabled = false;
        ruleColorLabelGroup.enabled = false;
        ruleColorRadioGroup.enabled = false;
        ruleHexGroup.enabled = false;
        applyPreviewFromUI();
    };
    rbWithRule.onClick = function() {
        rbNoRule.value = false;
        ruleInput.enabled = true;
        ruleColorLabelGroup.enabled = true;
        ruleColorRadioGroup.enabled = true;
        ruleHexGroup.enabled = true;
        applyPreviewFromUI();
    };
    ruleInput.onChange = applyPreviewFromUI;

    rbBlackLine.onClick = rbWhiteLine.onClick = function() {
        rbRuleHex.value = false;
        applyPreviewFromUI();
    };
    rbRuleHex.onClick = function() {
        ruleHexInput.enabled = true;
        rbBlackLine.value = false;
        rbWhiteLine.value = false;
        applyPreviewFromUI();
    };
    ruleHexInput.onChange = applyPreviewFromUI;

    // 右カラム: 書き出しサイズ
    var scalePanel = colRight.add("panel", undefined, LABELS.size[lang]);
    scalePanel.margins = [16, 20, 10, 10];
    scalePanel.orientation = "column";
    scalePanel.alignChildren = "left";

    // Set alignChildren to left for scalePanel
    scalePanel.alignChildren = "left";

    // 左カラム：倍率ラジオボタン（2カラム構造を廃止し、scalePanelに直接追加）
    var rbScale100 = scalePanel.add("radiobutton", undefined, label100);
    var rbScale200 = scalePanel.add("radiobutton", undefined, label200);
    var rbScale300 = scalePanel.add("radiobutton", undefined, label300);
    var rbScale400 = scalePanel.add("radiobutton", undefined, label400);
    rbScale400.value = true;

    // Store original detailed labels for scale radio buttons with pixel info
    var label100Orig = "1x：" + toPx(width) + " × " + toPx(height);
    var label200Orig = "2x：" + toPx(width * 2) + " × " + toPx(height * 2);
    var label300Orig = "3x：" + toPx(width * 3) + " × " + toPx(height * 3);
    var label400Orig = "4x：" + toPx(width * 4) + " × " + toPx(height * 4);

    // --- Helper functions for scale radio buttons ---
    function disableOtherScaleInputs() {
        if (typeof widthInput !== "undefined" && widthInput !== null) {
            if (typeof widthInput !== "undefined" && widthInput !== null) {
                widthInput.enabled = false;
            }
        }
        if (typeof rbWidthInput !== "undefined" && rbWidthInput !== null) {
            rbWidthInput.value = false;
        }
        if (typeof rbCustomScale !== "undefined" && rbCustomScale !== null) {
            rbCustomScale.value = false;
        }
        if (typeof customScaleInput !== "undefined" && customScaleInput !== null) {
            customScaleInput.enabled = false;
        }
    }

    function setCustomSuffix(value) {
        if (typeof rbNoSuffix !== "undefined" && rbNoSuffix !== null) {
            rbNoSuffix.value = false;
        }
        if (typeof rbCustomSuffix !== "undefined" && rbCustomSuffix !== null) {
            rbCustomSuffix.value = true;
        }
        if (typeof suffixInput !== "undefined" && suffixInput !== null) {
            suffixInput.enabled = true;
            suffixInput.text = value;
        }
        updateFilenamePreview();
    }

    function updateScaleButtonColors(active) {
        var gray = [0.5, 0.5, 0.5],
            black = [0, 0, 0];
        rbScale100.graphics.foregroundColor = rbScale100.graphics.newPen(rbScale100.graphics.PenType.SOLID_COLOR, active === "100" ? black : gray, 1);
        rbScale200.graphics.foregroundColor = rbScale200.graphics.newPen(rbScale200.graphics.PenType.SOLID_COLOR, active === "200" ? black : gray, 1);
        rbScale300.graphics.foregroundColor = rbScale300.graphics.newPen(rbScale300.graphics.PenType.SOLID_COLOR, active === "300" ? black : gray, 1);
        rbScale400.graphics.foregroundColor = rbScale400.graphics.newPen(rbScale400.graphics.PenType.SOLID_COLOR, active === "400" ? black : gray, 1);
    }

    // Refactored scale radio button handlers using a loop and handler function
    var scaleRadios = {
        "100": {
            button: rbScale100,
            label: label100Orig,
            multiplier: 1
        },
        "200": {
            button: rbScale200,
            label: label200Orig,
            multiplier: 2
        },
        "300": {
            button: rbScale300,
            label: label300Orig,
            multiplier: 3
        },
        "400": {
            button: rbScale400,
            label: label400Orig,
            multiplier: 4
        }
    };

    function createScaleClickHandler(scaleKey) {
        return function() {
            clearAllScaleRadioSelections();
            var config = scaleRadios[scaleKey];
            config.button.value = true;

            rbScale100.text = "1x：" + toPx(width * 1) + " × " + toPx(height * 1);
            rbScale200.text = "2x：" + toPx(width * 2) + " × " + toPx(height * 2);
            rbScale300.text = "3x：" + toPx(width * 3) + " × " + toPx(height * 3);
            rbScale400.text = "4x：" + toPx(width * 4) + " × " + toPx(height * 4);

            config.button.text = config.label;

            disableOtherScaleInputs();
            updateScaleButtonColors(scaleKey);
            setCustomSuffix(scaleKey);
            if (typeof customScaleInput !== "undefined" && customScaleInput !== null) {
                customScaleInput.text = scaleKey;
            }
            if (typeof widthInput !== "undefined" && widthInput !== null) {
                widthInput.text = String(toPx(width * config.multiplier));
            }
            updateFilenamePreview();
        };
    }

    // Assign the handlers
    rbScale100.onClick = createScaleClickHandler("100");
    rbScale200.onClick = createScaleClickHandler("200");
    rbScale300.onClick = createScaleClickHandler("300");
    rbScale400.onClick = createScaleClickHandler("400");
    // Optionally, initialize the text and color of rbScale200 to show its detailed label and selection state since it's selected by default
    rbScale100.text = label100Orig;
    rbScale200.text = label200Orig;
    rbScale300.text = label300Orig;
    rbScale400.text = label400Orig;
    // Set initial button colors for default selection (4x)
    rbScale100.graphics.foregroundColor = rbScale100.graphics.newPen(rbScale100.graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1);
    rbScale200.graphics.foregroundColor = rbScale200.graphics.newPen(rbScale200.graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1);
    rbScale300.graphics.foregroundColor = rbScale300.graphics.newPen(rbScale300.graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1);
    // Initialize color state and suffix for default (4x)
    if (typeof rbScale400.onClick === "function") rbScale400.onClick();

    // カスタム倍率ラジオボタンと入力欄の追加（倍率）
    var customScaleGroup = scalePanel.add("group");
    customScaleGroup.orientation = "row";
    var rbCustomScale = customScaleGroup.add("radiobutton", undefined, LABELS.customScaleLabel[lang]);
    rbCustomScale.preferredSize.width = (lang === "ja") ? 60 : 70;
    var customScaleInput = customScaleGroup.add("edittext", undefined, "500");
    customScaleInput.characters = 5;
    customScaleInput.enabled = false;
    var scaleLabel = customScaleGroup.add("statictext", undefined, "%");

    // --- 幅指定ラジオボタンと入力欄の追加（幅指定）
    var widthGroup = scalePanel.add("group");
    widthGroup.orientation = "row";
    var rbWidthInput = widthGroup.add("radiobutton", undefined, LABELS.widthInput[lang]);
    rbWidthInput.preferredSize.width = (lang === "ja") ? 60 : 70;
    var widthInput = widthGroup.add("edittext", undefined, "");
    widthInput.characters = 7;
    var widthLabel = widthGroup.add("statictext", undefined, "px");
    widthInput.enabled = false;

    // --- Set default widthInput based on default scale value (500) ---
    var defaultScale = 500;
    var scaledWidth = Math.round(width * defaultScale / 100);
    widthInput.text = scaledWidth.toString();

    // ラジオボタン選択時の入力有効化と排他制御
    rbWidthInput.onClick = function() {
        clearAllScaleRadioSelections();
        rbWidthInput.value = true;
        widthInput.enabled = true;
        customScaleInput.enabled = false;
    };
    rbCustomScale.onClick = function() {
        clearAllScaleRadioSelections();
        rbCustomScale.value = true;
        customScaleInput.enabled = true;
        widthInput.enabled = false;
    };

    // 右カラム: 書き出しファイル名パネル追加
    var namePanel = colRight.add("panel", undefined, LABELS.fileName[lang]);
    namePanel.margins = [16, 20, 10, 10];
    namePanel.orientation = "column";
    namePanel.alignChildren = "left";
    namePanel.preferredSize.width = 300;

    // ファイル名参照のオプション（UIのみ）
    var refNameGroup = namePanel.add("group");
    refNameGroup.orientation = "row";
    refNameGroup.add("statictext", undefined, LABELS.fileNameReference[lang]);
    var rbUseName = refNameGroup.add("radiobutton", undefined, "参照する");
    var rbNoUseName = refNameGroup.add("radiobutton", undefined, "参照しない");
    rbUseName.value = true;
    rbUseName.onClick = function() {
        updateFilenamePreview();
    };
    rbNoUseName.onClick = function() {
        rbCustomSuffix.value = true;
        suffixInput.enabled = true;
        suffixInput.active = true;
        rbNone.value = true;
        rbDash.value = false;
        rbUnderscore.value = false;
        updateFilenamePreview();
    };
    // --- 記号・接尾辞パネルの新レイアウト ---
    // 記号グループ（ラベル＋○なし ○- ○_）
    var symbolGroup = namePanel.add("group");
    symbolGroup.orientation = "row";
    symbolGroup.add("statictext", undefined, LABELS.symbol[lang]);
    var rbNone = symbolGroup.add("radiobutton", undefined, LABELS.none[lang]);
    var rbDash = symbolGroup.add("radiobutton", undefined, "-");
    var rbUnderscore = symbolGroup.add("radiobutton", undefined, "_");
    rbDash.value = true;

    // 接尾辞グループ（横並び）
    var suffixGroup = namePanel.add("group");
    suffixGroup.orientation = "row";
    suffixGroup.alignChildren = "left, center";
    suffixGroup.add("statictext", undefined, LABELS.suffixTextLabel[lang]);
    var rbNoSuffix = suffixGroup.add("radiobutton", undefined, LABELS.none[lang]);
    var rbCustomSuffix = suffixGroup.add("radiobutton", undefined, "");
    var suffixInput = suffixGroup.add("edittext", undefined, "");
    suffixInput.characters = 14;
    suffixInput.enabled = false;

    rbNoSuffix.value = true;
    rbCustomSuffix.value = false;

    var filenamePreviewGroup = namePanel.add("panel", undefined, "");
    filenamePreviewGroup.size = [300, 46];
    filenamePreviewGroup.orientation = "row";
    filenamePreviewGroup.alignChildren = "left";

    var filenamePreview = filenamePreviewGroup.add("statictext", undefined, "");
    filenamePreview.preferredSize = [260, 22]; // 高さを明示的に設定しディセンダーが見切れないように
    filenamePreview.justify = "left";

    rbNoSuffix.onClick = function() {
        suffixInput.enabled = false;
        updateFilenamePreview();
    };
    rbCustomSuffix.onClick = function() {
        suffixInput.enabled = true;
        updateFilenamePreview();
    };

    // --- 接尾辞をスケールや幅指定に連動して自動設定するロジックを追加 ---
    rbWidthInput.onClick = function() {
        clearAllScaleRadioSelections();
        rbWidthInput.value = true;
        widthInput.enabled = true;
        customScaleInput.enabled = false;
        setSuffixFrom(widthInput.text);
    };
    widthInput.onChange = function() {
        setSuffixFrom(widthInput.text);
    };
    rbCustomScale.onClick = function() {
        clearAllScaleRadioSelections();
        rbCustomScale.value = true;
        customScaleInput.enabled = true;
        widthInput.enabled = false;
        setSuffixFrom(customScaleInput.text);
    };
    customScaleInput.onChange = function() {
        setSuffixFrom(customScaleInput.text);
    };

    // 右カラム: 書き出し先（ラジオボタン横並び、ラベル調整）
    var locPanel = colRight.add("panel", undefined, LABELS.location[lang]);
    locPanel.margins = [16, 20, 10, 10];
    locPanel.orientation = "column";
    locPanel.alignChildren = "left";
    locPanel.preferredSize.width = 300;



    var locGroup = locPanel.add("group");
    locGroup.orientation = "row";
    locGroup.alignChildren = "left";
    var rbDesktop = locGroup.add("radiobutton", undefined, LABELS.desktop[lang]);
    var rbSameFolder = locGroup.add("radiobutton", undefined, LABELS.sameFolder[lang]);
    rbDesktop.value = true;

    // プリセット選択で値を自動入力
    presetDropdown.onChange = function() {
        var idx = presetDropdown.selection.index;
        if (idx === 0) return; // カスタム

        var p = presets[idx - 1];

        // 背景ラジオボタンと入力欄の状態をすべてリセット
        rbTransparent.value = false;
        rbWhite.value = false;
        rbBlack.value = false;
        rbHex.value = false;
        rbTransparentGrid.value = false;
        hexInput.enabled = false;
        transparentGridInput.enabled = false;

        // 背景値に応じて正しく反映
        if (p.background === "transparent") {
            rbTransparent.value = true;
        } else if (p.background === "white") {
            rbWhite.value = true;
        } else if (p.background === "black") {
            rbBlack.value = true;
        } else if (p.background === "transparentGrid") {
            rbTransparentGrid.value = true;
            transparentGridInput.enabled = true;
        } else if (typeof p.background === "string" && p.background.charAt(0) === "#") {
            rbHex.value = true;
            hexInput.enabled = true;
            hexInput.text = p.background;
        }

        if (p.margin === "none") {
            rbNoMargin.value = true;
            marginInput.enabled = false;
        } else if (p.margin.indexOf("horizontal:") === 0) {
            rbMarginH.value = true;
            marginInput.text = p.margin.split(":")[1];
            marginInput.enabled = true;
        } else if (p.margin.indexOf("vertical:") === 0) {
            rbMarginV.value = true;
            marginInput.text = p.margin.split(":")[1];
            marginInput.enabled = true;
        } else if (p.margin.indexOf("all:") === 0) {
            rbMarginAll.value = true;
            marginInput.text = p.margin.split(":")[1];
            marginInput.enabled = true;
        } else {
            rbMarginAll.value = true;
            marginInput.text = p.margin;
            marginInput.enabled = true;
        }

        if (p.rule === "none") {
            rbNoRule.value = true;
            rbWithRule.value = false;
            ruleInput.enabled = false;
            ruleColorLabelGroup.enabled = false;
            ruleColorRadioGroup.enabled = false;
            ruleHexGroup.enabled = false;
        } else {
            rbWithRule.value = true;
            rbNoRule.value = false;
            var parts = p.rule.split(",");
            ruleInput.text = parts[0];
            ruleInput.enabled = true;
            ruleColorLabelGroup.enabled = true;
            ruleColorRadioGroup.enabled = true;
            ruleHexGroup.enabled = true;
            rbBlackLine.value = (parts[1] === "black");
            rbWhiteLine.value = (parts[1] === "white");
        }

        rbDesktop.value = (p.location === "desktop");
        rbSameFolder.value = (p.location !== "desktop");

        // 記号選択の復元
        rbNone.value = (p.symbol === "");
        rbDash.value = (p.symbol === "-");
        rbUnderscore.value = (p.symbol === "_");

        // 接尾辞選択の復元
        if (p.suffix === "") {
            rbNoSuffix.value = true;
            rbCustomSuffix.value = false;
            suffixInput.enabled = false;
        } else {
            rbNoSuffix.value = false;
            rbCustomSuffix.value = true;
            suffixInput.enabled = true;
            suffixInput.text = p.suffix;
        }

        // --- 書き出しサイズ（scale or width）を正しく反映 ---
        rbScale100.value = false;
        rbScale200.value = false;
        rbScale300.value = false;
        rbWidthInput.value = false;
        rbCustomScale.value = false;
        customScaleInput.enabled = false;
        widthInput.enabled = false;

        if (p.size.indexOf("scale:") === 0) {
            var scaleValue = p.size.split(":")[1];
            if (scaleValue === "100") {
                rbScale100.value = true;
                if (typeof rbScale100.onClick === "function") rbScale100.onClick();
            } else if (scaleValue === "200") {
                rbScale200.value = true;
                if (typeof rbScale200.onClick === "function") rbScale200.onClick();
            } else if (scaleValue === "300") {
                rbScale300.value = true;
                if (typeof rbScale300.onClick === "function") rbScale300.onClick();
            } else if (scaleValue === "400") {
                rbScale400.value = true;
                if (typeof rbScale400.onClick === "function") rbScale400.onClick();
            } else {
                rbCustomScale.value = true;
                customScaleInput.enabled = true;
                customScaleInput.text = scaleValue;
                if (typeof rbCustomScale.onClick === "function") rbCustomScale.onClick();
            }
        } else if (p.size.indexOf("width:") === 0) {
            var widthValue = p.size.split(":")[1];
            rbWidthInput.value = true;
            widthInput.enabled = true;
            widthInput.text = widthValue;
            if (typeof rbWidthInput.onClick === "function") rbWidthInput.onClick();
        }
    };

    // --- Add checkbox for showing folder after export ---
    var showFolderGroup = locPanel.add("group");
    showFolderGroup.orientation = "row";
    var showFolderCheckbox = showFolderGroup.add("checkbox", undefined, LABELS.showFolder[lang]);
    showFolderCheckbox.value = true;
    if (Folder.fs !== "Macintosh") {
        showFolderGroup.visible = false;
    }



    // --- Filename preview update function (refactored) ---
    function updateFilenamePreview() {
        if (typeof filenamePreview === "undefined") return;

        var sym = "";
        if (rbDash && rbDash.value) sym = "-";
        else if (rbUnderscore && rbUnderscore.value) sym = "_";
        else sym = "";

        var suffix = rbNoSuffix && rbNoSuffix.value ? "" : suffixInput.text;
        var useDocName = rbNoUseName && rbNoUseName.value ? false : true;
        var docNameUsed = (typeof docName !== "undefined") ? docName : "Untitled";

        filenamePreview.text = generateExportFilename(docNameUsed, sym, suffix, useDocName);
    }

    // --- Symbol radio button event handlers for filename preview ---
    rbDash.onClick = function() {
        updateFilenamePreview();
    };
    rbUnderscore.onClick = function() {
        updateFilenamePreview();
    };
    rbNone.onClick = function() {
        updateFilenamePreview();
    };
    suffixInput.onChange = function() {
        updateFilenamePreview();
    };

    // (Do not call updateFilenamePreview here; call only in dialog "show" event)

    // OK/キャンセルボタン: 左端にキャンセル、右端にOK
    var btnGroup = dialog.add("group");
    btnGroup.alignment = "fill";
    btnGroup.alignChildren = ["right", "center"];
    var btnCancel = btnGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var btnOK = btnGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });
    // EnterキーでOKボタンを実行可能に
    dialog.defaultElement = btnOK;

    // --- プレビュー描画ロジックを関数として定義 ---
    function previewApply() {
        var allItems = app.activeDocument.pageItems;
        var selectedItems = app.activeDocument.selection;
        previewHiddenItems.length = 0; // clear existing before reuse

        function isInSelection(item, selectionArray) {
            while (item != null) {
                for (var i = 0; i < selectionArray.length; i++) {
                    if (item === selectionArray[i]) return true;
                }
                item = item.parent;
            }
            return false;
        }

        // 選択外オブジェクトを一時的に非表示
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            if (item.locked || item.layer.locked) continue;
            if (!isInSelection(item, selectedItems)) {
                // Robust check for hiding items
                if (
                    item &&
                    !item.hidden &&
                    item.editable !== false &&
                    !item.locked &&
                    item.layer &&
                    !item.layer.locked &&
                    (!item.parent || !item.parent.locked)
                ) {
                    try {
                        item.hidden = true;
                        previewHiddenItems.push(item);
                    } catch (e) {
                        alert("Error hiding item during previewApply: " + e.message);
                    }
                }
            }
        }

        // プレビュー用図形を削除
        removePreviewItems();

        var bg = "white";
        // 背景色選択状態を取得
        if (rbTransparent.value) {
            bg = "transparent";
        } else if (rbBlack.value) {
            bg = "black";
        } else if (rbHex.value) {
            bg = hexInput.text;
            if (bg.charAt(0) !== "#") bg = "#" + bg;
        } else if (rbTransparentGrid.value) {
            bg = "transparentGrid";
        }

        // --- Margin calculation (horizontal/vertical/all) ---
        var marginH = 0;
        var marginV = 0;
        var raw = parseFloat(marginInput.text);
        var factor = getRulerUnitInfo().factor;
        if (rbMarginH.value) {
            marginH = isNaN(raw) ? 0 : raw * factor;
        } else if (rbMarginV.value) {
            marginV = isNaN(raw) ? 0 : raw * factor;
        } else if (rbMarginAll.value) {
            var val = isNaN(raw) ? 0 : raw * factor;
            marginH = val;
            marginV = val;
        }

        var bounds = selectionBounds;
        var left = bounds[0] - marginH;
        var top = bounds[1] + marginV;
        var right = bounds[2] + marginH;
        var bottom = bounds[3] - marginV;

        left = Math.floor(left);
        top = Math.ceil(top);
        right = Math.ceil(right);
        bottom = Math.floor(bottom);

        var width = right - left;
        var height = top - bottom;

        // --- 背景描画処理を共通関数で ---
        var tilePercent = parseFloat(transparentGridInput.text);
        if (isNaN(tilePercent) || tilePercent <= 0) tilePercent = 100;
        createExportBackground(bg, left, top, width, height, tilePercent);

        // --- 罫線描画（最小1px保証、整数化） ---
        var ruleWidthRaw = rbWithRule.value ? parseFloat(ruleInput.text) : 0;
        var unitFactor = getRulerUnitInfo().factor;
        var ruleWidth = isNaN(ruleWidthRaw) ? 0 : ruleWidthRaw * unitFactor;

        if (ruleWidth > 0) {
            ruleWidth = Math.ceil(ruleWidth); // pxに変換後、切り上げで整数化
            var adjustedStrokeWidth = Math.max(1, ruleWidth); // 最低1pxを保証、奇数でもそのまま
            var halfStroke = adjustedStrokeWidth / 2;
            var innerLeft = left + halfStroke;
            var innerTop = top - halfStroke;
            var innerWidth = width - adjustedStrokeWidth;
            var innerHeight = height - adjustedStrokeWidth;

            var ruleRect = app.activeDocument.pathItems.rectangle(innerTop, innerLeft, innerWidth, innerHeight);
            ruleRect.name = "preview_border";
            ruleRect.stroked = true;
            if (rbBlackLine.value) {
                ruleRect.strokeColor = createBlackColor();
            } else if (rbWhiteLine.value) {
                ruleRect.strokeColor = createWhiteColor();
            } else if (rbRuleHex.value) {
                var ruleColorChoice = ruleHexInput.text;
                if (ruleColorChoice.charAt(0) !== "#") ruleColorChoice = "#" + ruleColorChoice;
                ruleRect.strokeColor = createColorFromCode(ruleColorChoice);
            }
            ruleRect.strokeWidth = adjustedStrokeWidth;
            ruleRect.filled = false;
        }

        app.redraw();
    }

    // --- Store deep copy of current hidden states before dialog shows ---
    var doc = app.activeDocument;
    var originalHiddenStates = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        originalHiddenStates.push({
            item: item,
            hidden: item.hidden
        });
    }

    applyPreviewFromUI();

    // --- ここから追加: ダイアログ表示時のイベントでファイル名プレビュー更新とセンタリング＆位置調整 ---
    dialog.addEventListener("show", function() {
        updateFilenamePreview();
        dialog.center();
        dialog.location = [dialog.location[0] + 300, dialog.location[1]];
    });

    var dialogResult = dialog.show();

    // プレビュー用の仮図形削除
    removePreviewItems();

    // 一時的に非表示にしたオブジェクトと元のhidden状態を一括復元
    restoreAllHiddenStates(temporaryHiddenItems, originalHiddenStates);

    if (dialogResult === 1) {
        var bg = "white";
        if (rbTransparent.value) {
            bg = "transparent";
        } else if (rbBlack.value) {
            bg = "black";
        } else if (rbHex.value) {
            bg = hexInput.text;
            if (bg.charAt(0) !== "#") bg = "#" + bg;
        } else if (rbTransparentGrid.value) {
            bg = "transparentGrid";
        }

        var location = rbDesktop.value ? "desktop" : "same";
        var scaleValue = getScaleValueFromUI();
        var widthOverride = null;

        if (rbWidthInput.value) {
            var val = parseFloat(widthInput.text);
            if (!isNaN(val) && val > 0) {
                widthOverride = val;
                scaleValue = null;
            }
        } else if (rbCustomScale.value) {
            var val = parseFloat(customScaleInput.text);
            if (!isNaN(val) && val > 0) {
                scaleValue = val;
            }
        }

        // --- marginValue: use string encoding as per new spec ---
        var marginValue = "none";
        if (rbMarginH.value) {
            marginValue = "horizontal:" + marginInput.text;
        } else if (rbMarginV.value) {
            marginValue = "vertical:" + marginInput.text;
        } else if (rbMarginAll.value) {
            marginValue = "all:" + marginInput.text;
        }

        var ruleWidthRaw = rbWithRule.value ? parseFloat(ruleInput.text) : 0;
        var ruleWidth = isNaN(ruleWidthRaw) ? 0 : ruleWidthRaw * unitFactor;
        var ruleColorChoice;
        if (rbBlackLine.value) {
            ruleColorChoice = "black";
        } else if (rbWhiteLine.value) {
            ruleColorChoice = "white";
        } else if (rbRuleHex.value) {
            ruleColorChoice = ruleHexInput.text;
            if (ruleColorChoice.charAt(0) !== "#") ruleColorChoice = "#" + ruleColorChoice;
        }
        var symbol = rbDash.value ? "-" : rbUnderscore.value ? "_" : "";
        var suffix = (rbNoSuffix.value) ? "" : suffixInput.text;

        // For transparentGrid, pass tile percent as a property for main()
        var transparentGridPercent = null;
        if (bg === "transparentGrid") {
            transparentGridPercent = parseFloat(transparentGridInput.text);
            if (isNaN(transparentGridPercent) || transparentGridPercent <= 0) transparentGridPercent = 100;
        }

        // Add previewHiddenItems to return value
        return {
            bgChoice: bg,
            outputLocation: location,
            scaleValue: scaleValue,
            widthOverride: widthOverride,
            marginValue: marginValue,
            ruleWidth: ruleWidth,
            ruleColor: ruleColorChoice,
            suffix: suffix,
            symbol: symbol,
            showFolder: showFolderCheckbox.value,
            transparentGridPercent: transparentGridPercent,
            exportFileName: filenamePreview.text,
            hiddenItems: previewHiddenItems // newly returned list
        };
    } else {
        // キャンセル時も temporaryHiddenItems を再表示
        // (already done above)
        return null;
    }
    // -------------------------------
    // Helper: Get scale value from UI
    // -------------------------------
    function getScaleValueFromUI() {
        if (rbScale100 && rbScale100.value) return 100;
        if (rbScale200 && rbScale200.value) return 200;
        if (rbScale300 && rbScale300.value) return 300;
        if (rbScale400 && rbScale400.value) return 400;
        return null;
    }
}


main();

function main() {
    var rulerUnitInfo = getRulerUnitInfo();

    if (app.documents.length === 0 || app.selection.length === 0) {
        alert(LABELS.noSelection[lang]);
        return;
    }

    var doc = app.activeDocument;
    docName = doc.name.replace(/\.ai$/i, "");
    var previewLayer = doc.layers.add();
    previewLayer.name = PREVIEW_LAYER_NAME;

    var originalSelection = doc.selection;
    var selectionCopy = duplicateSelectionToLayer(originalSelection, previewLayer);

    var hiddenLayers = hideOtherLayers(doc, previewLayer);

    var selectedItems = doc.selection;
    var bounds = getSelectionBounds(selectedItems);

    // すべての pageItems の hidden 状態を保存
    var originalHiddenStates = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
        originalHiddenStates.push({
            item: doc.pageItems[i],
            hidden: doc.pageItems[i].hidden
        });
    }
    // アートボードのアクティブ状態も保存
    var originalABIndex = doc.artboards.getActiveArtboardIndex();

    // showExportOptionsDialog will now return an object with hiddenItems, extract accordingly
    var dialogResult = showExportOptionsDialog(bounds);
    var userChoice = dialogResult;
    var temporaryHiddenItems = (dialogResult && dialogResult.hiddenItems) ? dialogResult.hiddenItems : [];
    // キャンセル時にも hidden 状態とアートボード状態を復元
    if (!userChoice) {
        restoreAllHiddenStates(temporaryHiddenItems, originalHiddenStates);
        // 非表示レイヤーの復元
        restoreLayerVisibility(hiddenLayers);
        // preview_background 削除（キャンセル時にも明示的に削除）
        removePreviewItems();
        doc.artboards.setActiveArtboardIndex(originalABIndex);
        // プレビュー用レイヤーを削除
        removePreviewLayerByName(doc, PREVIEW_LAYER_NAME);
        return;
    }
    var bgChoice = userChoice.bgChoice;
    var scaleValue = userChoice.scaleValue;
    var widthOverride = userChoice.widthOverride;
    // --- marginValue 分解: marginH, marginV ---
    var marginH = 0;
    var marginV = 0;
    if (typeof userChoice.marginValue === "string") {
        if (userChoice.marginValue.indexOf("horizontal:") === 0) {
            var raw = parseFloat(userChoice.marginValue.split(":")[1]);
            marginH = isNaN(raw) ? 0 : raw * getRulerUnitInfo().factor;
        } else if (userChoice.marginValue.indexOf("vertical:") === 0) {
            var raw = parseFloat(userChoice.marginValue.split(":")[1]);
            marginV = isNaN(raw) ? 0 : raw * getRulerUnitInfo().factor;
        } else if (userChoice.marginValue.indexOf("all:") === 0) {
            var raw = parseFloat(userChoice.marginValue.split(":")[1]);
            var val = isNaN(raw) ? 0 : raw * getRulerUnitInfo().factor;
            marginH = val;
            marginV = val;
        }
    } else if (typeof userChoice.marginValue === "number") {
        marginH = marginV = userChoice.marginValue;
    }
    var ruleWidth = userChoice.ruleWidth;
    var ruleColor = userChoice.ruleColor;
    var suffix = userChoice.suffix;
    var symbol = userChoice.symbol;

    // --- ruleWidth整数化とadjustedStrokeWidth計算 ---
    ruleWidth = Math.ceil(ruleWidth); // pxに変換後、切り上げで整数化
    var adjustedStrokeWidth = Math.max(1, ruleWidth); // 最低1pxを保証、奇数でもそのまま

    // 書き出し範囲（選択オブジェクトのvisibleBounds, marginH/marginVを使用）
    var left = bounds[0] - marginH;
    var top = bounds[1] + marginV;
    var right = bounds[2] + marginH;
    var bottom = bounds[3] - marginV;
    // ピクセルパーフェクトな整数に丸める
    left = Math.floor(left);
    top = Math.ceil(top);
    right = Math.ceil(right);
    bottom = Math.floor(bottom);

    var width = right - left;
    var height = top - bottom;
    if (width <= 0 || height <= 0) {
        alert(LABELS.invalidSize[lang]);
        // hidden 状態を復元
        for (var i = 0; i < originalHiddenStates.length; i++) {
            try {
                var obj = originalHiddenStates[i].item;
                if (obj && obj.isValid) {
                    obj.hidden = originalHiddenStates[i].hidden;
                }
            } catch (e) {
                // ログ: 無効サイズ時hidden復元エラー
                alert("Error restoring hidden state after invalid size in main: " + e.message);
            }
        }
        return;
    }

    // 一時アートボードの作成
    var originalABCount = doc.artboards.length;
    var newAB = doc.artboards.add([left, top, right, bottom]);
    var newABIndex = doc.artboards.length - 1;
    doc.artboards.setActiveArtboardIndex(newABIndex);

    // Exportオプションを先に定義
    var exportOptions = new ExportOptionsPNG24();
    exportOptions.transparency = (bgChoice === "transparent");
    exportOptions.artBoardClipping = true;

    // 背景描画処理を共通関数で
    var tilePercent = (typeof userChoice.transparentGridPercent === "number") ? userChoice.transparentGridPercent : 100;
    var whiteRect = createExportBackground(bgChoice, left, top, width, height, tilePercent);
    exportOptions.transparency = (bgChoice === "transparent") ? true : false;

    // 罫線描画（白背景の上）
    if (ruleWidth > 0) {
        var halfStroke = adjustedStrokeWidth / 2;
        var innerLeft = left + halfStroke;
        var innerTop = top - halfStroke;
        var innerWidth = width - adjustedStrokeWidth;
        var innerHeight = height - adjustedStrokeWidth;

        var ruleRect = doc.pathItems.rectangle(innerTop, innerLeft, innerWidth, innerHeight);
        ruleRect.stroked = true;
        if (ruleColor === "black") {
            ruleRect.strokeColor = createBlackColor();
        } else if (ruleColor === "white") {
            ruleRect.strokeColor = createWhiteColor();
        } else if (typeof ruleColor === "string" && ruleColor.charAt(0) === "#" && ruleColor.length === 7) {
            ruleRect.strokeColor = createColorFromCode(ruleColor);
        }
        ruleRect.strokeWidth = adjustedStrokeWidth;
        ruleRect.filled = false;
        ruleRect.zOrder(ZOrderMethod.BRINGTOFRONT);
    }

    // 書き出しパス設定
    var docFolder = (userChoice.outputLocation === "desktop") ?
        Folder.desktop :
        doc.fullName.parent;
    // Always use the filename preview shown to the user
    var fileName = userChoice.exportFileName;
    var exportFile = new File(docFolder + "/" + fileName);

    // スケール倍率をエクスポート直前に設定
    var scaleRatio;
    if (widthOverride !== null) {
        var actualPxWidth = width / rulerUnitInfo.factor;
        scaleRatio = (widthOverride / actualPxWidth) * 100;
    } else {
        scaleRatio = parseFloat(scaleValue);
    }

    if (isNaN(scaleRatio) || scaleRatio <= 0) {
        scaleRatio = 100; // fallback
    }
    exportOptions.horizontalScale = scaleRatio;
    exportOptions.verticalScale = scaleRatio;

    try {
        doc.exportFile(exportFile, ExportType.PNG24, exportOptions);
    } catch (e) {
        alert(LABELS.exportFailed[lang] + String.fromCharCode(13) + e.message);
        alert("Error during exportFile: " + e.message);
    }

    // 背景オブジェクトの削除を共通関数で一元化
    removeExportBackground(bgChoice);
    // 罫線矩形の削除（背景とは別管理）
    if (typeof ruleRect !== "undefined") {
        try {
            ruleRect.remove();
        } catch (e) {
            // ログ: 罫線矩形削除時エラー
            alert("Error removing ruleRect after export: " + e.message);
        }
    }

    // OK押下後にも一時的に非表示にしたものとhidden状態を正確に復元
    restoreAllHiddenStates(temporaryHiddenItems, originalHiddenStates);

    // redraw and cleanup
    app.redraw();
    removePreviewItems();

    // 一時アートボードの削除と復元
    try {
        doc.artboards.remove(newABIndex);
    } catch (e) {
        // ログ: 一時アートボード削除時エラー
        alert("Error removing temporary artboard: " + e.message);
    }
    doc.artboards.setActiveArtboardIndex(originalABIndex);

    // プレビュー用レイヤーを削除
    removePreviewLayerByName(doc, PREVIEW_LAYER_NAME);

    // 元のレイヤーの再表示（hiddenLayers に記録されたレイヤーを復元）
    restoreLayerVisibility(hiddenLayers);

    if (Folder.fs === "Macintosh" && userChoice.showFolder) {
        docFolder.execute();
    }

    // 念のためプレビュー用レイヤーを再削除（終了時の最終クリーンアップ）
    removePreviewLayerByName(app.activeDocument, PREVIEW_LAYER_NAME);
}

// 選択オブジェクトのバウンディングボックスを取得
function getSelectionBounds(selection) {
    var left = Number.POSITIVE_INFINITY;
    var top = Number.NEGATIVE_INFINITY;
    var right = Number.NEGATIVE_INFINITY;
    var bottom = Number.POSITIVE_INFINITY;

    for (var i = 0; i < selection.length; i++) {
        var bounds = selection[i].visibleBounds;
        if (bounds[0] < left) left = bounds[0];
        if (bounds[1] > top) top = bounds[1];
        if (bounds[2] > right) right = bounds[2];
        if (bounds[3] < bottom) bottom = bounds[3];
    }

    return [left, top, right, bottom];
}

// 白色カラーを生成
function createWhiteColor() {
    if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
        var white = new RGBColor();
        white.red = 255;
        white.green = 255;
        white.blue = 255;
        return white;
    } else {
        var white = new CMYKColor();
        white.cyan = 0;
        white.magenta = 0;
        white.yellow = 0;
        white.black = 0;
        return white;
    }
}

function createBlackColor() {
    if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
        var black = new RGBColor();
        black.red = 0;
        black.green = 0;
        black.blue = 0;
        return black;
    } else {
        var black = new CMYKColor();
        black.cyan = 0;
        black.magenta = 0;
        black.yellow = 0;
        black.black = 100;
        return black;
    }
}

function getRulerUnitInfo() {
    var t = app.preferences.getIntegerPreference("rulerType");
    var units = {
        label: "pt",
        factor: 1.0
    };
    if (t === 0) units = {
        label: "inch",
        factor: 72.0
    };
    else if (t === 1) units = {
        label: "mm",
        factor: 72.0 / 25.4
    };
    else if (t === 3) units = {
        label: "pica",
        factor: 12.0
    };
    else if (t === 4) units = {
        label: "cm",
        factor: 72.0 / 2.54
    };
    else if (t === 5) units = {
        label: "Q",
        factor: 72.0 / 25.4 * 0.25
    };
    else if (t === 6) units = {
        label: "px",
        factor: 1.0
    };
    return units;
}

function createColorFromCode(value) {
    value = value.replace(/^\s+|\s+$/g, "").toUpperCase().replace(/\s+/g, "");

    // Hex (#RRGGBB)
    if (/^#[0-9A-F]{6}$/.test(value)) {
        var rgb = new RGBColor();
        rgb.red = parseInt(value.substring(1, 3), 16);
        rgb.green = parseInt(value.substring(3, 5), 16);
        rgb.blue = parseInt(value.substring(5, 7), 16);
        return rgb;
    }

    // RGB (R255G255B255)
    var rgbMatch = value.match(/^R(\d{1,3})G(\d{1,3})B(\d{1,3})$/i);
    if (rgbMatch) {
        var rgb = new RGBColor();
        rgb.red = Math.min(255, parseInt(rgbMatch[1], 10));
        rgb.green = Math.min(255, parseInt(rgbMatch[2], 10));
        rgb.blue = Math.min(255, parseInt(rgbMatch[3], 10));
        return rgb;
    }

    // CMYK (C0M100Y100K0)
    var cmykMatch = value.match(/^C(\d{1,3})M(\d{1,3})Y(\d{1,3})K(\d{1,3})$/i);
    if (cmykMatch) {
        var cmyk = new CMYKColor();
        cmyk.cyan = Math.min(100, parseInt(cmykMatch[1], 10));
        cmyk.magenta = Math.min(100, parseInt(cmykMatch[2], 10));
        cmyk.yellow = Math.min(100, parseInt(cmykMatch[3], 10));
        cmyk.black = Math.min(100, parseInt(cmykMatch[4], 10));
        return cmyk;
    }

    return null;
}
// プレビュー用のアイテムを削除するユーティリティ
function removePreviewItems() {
    var doc = app.activeDocument;
    var itemsToRemove = [];

    // 検索対象を全 pageItems に拡張
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        if (item.name === "preview_background" || item.name === "preview_border") {
            itemsToRemove.push(item);
        }
    }

    // 削除処理（try-catch）
    for (var j = 0; j < itemsToRemove.length; j++) {
        try {
            if (itemsToRemove[j] && itemsToRemove[j].isValid) {
                itemsToRemove[j].remove();
            }
        } catch (e) {
            alert("Error removing preview item: " + e.message);
        }
    }
}

// チェッカーグリッド背景描画（透明グリッド用）
// parentGroup: 追加先グループ、left/top: 左上座標、width/height: サイズ、tileSize: タイル1マスの大きさ
function drawCheckerPattern(parentGroup, left, top, width, height, tileSize) {
    var doc = app.activeDocument;
    var isRGB = (doc.documentColorSpace === DocumentColorSpace.RGB);

    var grayColor = isRGB ? new RGBColor() : new CMYKColor();
    var whiteColor = isRGB ? new RGBColor() : new CMYKColor();

    if (isRGB) {
        grayColor.red = 204;
        grayColor.green = 204;
        grayColor.blue = 204;
        whiteColor.red = 255;
        whiteColor.green = 255;
        whiteColor.blue = 255;
    } else {
        grayColor.cyan = 0;
        grayColor.magenta = 0;
        grayColor.yellow = 0;
        grayColor.black = 30;
        whiteColor.cyan = 0;
        whiteColor.magenta = 0;
        whiteColor.yellow = 0;
        whiteColor.black = 0;
    }

    var cols = Math.ceil(width / tileSize);
    var rows = Math.ceil(height / tileSize);

    // チェックパターンを2色交互に敷き詰め
    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {
            var color = ((row + col) % 2 === 0) ? grayColor : whiteColor;
            var x = left + (col * tileSize);
            var y = top - (row * tileSize);

            var rect = parentGroup.pathItems.rectangle(y, x, tileSize, tileSize);
            rect.filled = true;
            rect.stroked = false;
            rect.fillColor = color;
        }
    }
    // 背面へ
    parentGroup.zOrder(ZOrderMethod.SENDTOBACK);
}
// -------------------------------
// 背景描画共通関数：書き出し・プレビューの両方で使用
// -------------------------------
function createExportBackground(bgType, left, top, width, height, tilePercent) {
    var doc = app.activeDocument;
    if (bgType === "transparentGrid") {
        var baseTile = 10 * getRulerUnitInfo().factor;
        var tileSize = baseTile * (tilePercent / 100);
        var patternGroup = doc.groupItems.add();
        patternGroup.name = "preview_background";
        drawCheckerPattern(patternGroup, left, top, width, height, tileSize);
        return patternGroup;
    } else if (bgType === "white" || bgType === "black" || (bgType.charAt(0) === "#" && bgType.length === 7)) {
        var rect = doc.pathItems.rectangle(top, left, width, height);
        rect.name = "preview_background";
        rect.filled = true;
        rect.stroked = false;

        if (bgType === "white") rect.fillColor = createWhiteColor();
        else if (bgType === "black") rect.fillColor = createBlackColor();
        else rect.fillColor = createColorFromCode(bgType);

        rect.zOrder(ZOrderMethod.SENDTOBACK);
        return rect;
    }
    return null;
}

// 背景オブジェクト（whiteRect または transparentGrid のグループ）を削除する関数
// 背景削除共通関数：白／黒／カラー／透明グリッドすべてに対応
function removeExportBackground(bgChoice) {
    var doc = app.activeDocument;
    try {
        if (bgChoice === "white" || bgChoice === "black" || (bgChoice.charAt(0) === "#" && bgChoice.length === 7)) {
            var whiteItem = doc.pageItems.getByName("preview_background");
            if (whiteItem && whiteItem.isValid) {
                whiteItem.remove();
            }
        } else if (bgChoice === "transparentGrid") {
            var patternGroup = doc.groupItems.getByName("preview_background");
            if (patternGroup && patternGroup.isValid) {
                patternGroup.remove();
            }
        }
    } catch (e) {
        alert("Error removing export background: " + e.message);
    }
}

// -------------------------------
// ファイル名の禁則文字・空白類を置換（記号に応じて "-" or "_"）
// -------------------------------
function sanitizeFilename(name, symbol) {
    var rep = (symbol === "-") ? "-" : (symbol === "_") ? "_" : "_";
    // alert("sanitizeFilename input name: " + name + ", symbol: " + symbol);
    return name.replace(/[¥\/:*?"<>|\r\n\t　 ]/g, rep);
}

// 強制的にプレビュー用レイヤーを削除する関数
// Remove preview layer by name, but do not alert if not found
function removePreviewLayerByName(doc, layerName) {
    try {
        var found = false;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                found = true;
                var targetLayer = doc.layers[i];
                if (targetLayer.locked) targetLayer.locked = false;
                if (!targetLayer.visible) targetLayer.visible = true;

                for (var j = targetLayer.pageItems.length - 1; j >= 0; j--) {
                    try {
                        targetLayer.pageItems[j].remove();
                    } catch (e) {}
                }

                targetLayer.remove();
                break;
            }
        }
        // silently skip if not found
    } catch (e) {
        alert("Error removing preview layer: " + e.message);
    }
}

// -------------------------------
// ファイル名生成関数（プレビューと書き出し共通化用）
// -------------------------------
function generateExportFilename(docName, symbol, suffix, useDocName) {
    var fileName = "";

    if (!suffix) {
        fileName = useDocName ? docName + "_selection.png" : "_selection.png";
    } else {
        fileName = (useDocName ? docName : "") + (symbol !== "" ? symbol : "") + suffix + ".png";
    }

    return sanitizeFilename(fileName, symbol);
}
// --- 元のhidden状態と一時的hidden項目をまとめて復元するヘルパー ---
function restoreAllHiddenStates(temporaryHiddenItems, originalHiddenStates) {
    // 一時的に非表示にしたものを再表示
    if (temporaryHiddenItems && temporaryHiddenItems.length) {
        restoreTemporaryHiddenItems(temporaryHiddenItems);
    }
    // 元のhidden状態を復元
    if (originalHiddenStates && originalHiddenStates.length) {
        for (var i = 0; i < originalHiddenStates.length; i++) {
            try {
                var obj = originalHiddenStates[i].item;
                if (obj && obj.isValid) {
                    obj.hidden = originalHiddenStates[i].hidden;
                }
            } catch (e) {
                alert("Error restoring hidden state: " + e.message);
            }
        }
    }
}