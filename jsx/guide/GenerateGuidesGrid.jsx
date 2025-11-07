#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

GenerateGuidesGrid.jsx

### 概要

- Illustrator のアートボードを指定した行数・列数に分割し、自動でガイドを生成するスクリプトです。
- 必要に応じてセルごとの長方形や裁ち落としガイドを描画でき、プリセットの書き出しにも対応します。
- すべての入力変更（キーボード/矢印キーを含む）でプレビューが即時更新

### 主な機能

- 行数・列数、ガター、上下左右マージン、ガイド伸張距離を自由に設定
- セルごとの長方形描画（不透明度15%）
- 裁ち落としガイドの追加描画
- 現在設定をプリセットとして書き出し
- すべてのアートボードに一括適用可能
- 日本語／英語インターフェース自動切替

### 処理の流れ

1. ダイアログで行数、列数、マージンなどを設定
2. プレビューを確認
3. OK をクリックしてガイド・長方形を作成

### オリジナル、謝辞

スガサワ君β  
https://note.com/sgswkn/n/nee8c3ec1a14c

### 更新履歴

- v1.3 (20251106) : すべてのUI入力（矢印キー含む）でリアルタイムプレビュー更新
- v1.2 (20250716) : 矢印キーでの数値増減機能追加
- v1.1 (20250427) : ガイド、裁ち落としガイド、プリセット書き出し機能追加
- v1.0 (20250424) : 初期バージョン

---

### Script Name:

GenerateGuidesGrid.jsx

### Overview

- A script for Illustrator that divides artboards into specified rows and columns, and automatically generates guides.
- Optionally draws cell rectangles and bleed guides, and supports exporting presets.
- Live preview updates on every input change (including keyboard/arrow keys)

### Main Features

- Freely set rows, columns, gutter, margins, and guide extension
- Draw cell rectangles (15% opacity)
- Add optional bleed guides
- Export current settings as presets
- Apply to all artboards at once
- Automatic Japanese / English UI switching

### Process Flow

1. Configure rows, columns, margins, etc. in the dialog
2. Check the preview
3. Click OK to create guides and rectangles

### Original / Acknowledgements

Sugasawa-kun β  
https://note.com/sgswkn/n/nee8c3ec1a14c

### Update History

- v1.3 (20251106): Live preview refresh on all UI changes (including arrow keys)
- v1.2 (20250716): Added arrow key value increment feature
- v1.1 (20250427): Added guides, bleed guides, and preset export feature
- v1.0 (20250424): Initial version

*/

var SCRIPT_VERSION = "v1.3";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

// ラベル定義 / Label definitions
var LABELS = {
    dialogTitle: {
        ja: "グリッドに分割 Pro " + SCRIPT_VERSION,
        en: "Split into Grid Pro " + SCRIPT_VERSION
    },
    presetLabel: {
        ja: "プリセット：",
        en: "Preset:"
    },
    rowTitle: {
        ja: "行（─ 横線）",
        en: "Rows (─ Horizontal)"
    },
    columnTitle: {
        ja: "列（│ 縦線）",
        en: "Columns (│ Vertical)"
    },
    rowsLabel: {
        ja: "行数：",
        en: "Rows:"
    },
    rowGutterLabel: {
        ja: "行間",
        en: "Gutter"
    },
    columnsLabel: {
        ja: "列数：",
        en: "Columns:"
    },
    colGutterLabel: {
        ja: "列間",
        en: "Gutter"
    },
    marginTitle: {
        ja: "マージン設定",
        en: "Margin Settings"
    },
    topLabel: {
        ja: "上：",
        en: "Top:"
    },
    leftLabel: {
        ja: "左：",
        en: "Left:"
    },
    bottomLabel: {
        ja: "下：",
        en: "Bottom:"
    },
    rightLabel: {
        ja: "右：",
        en: "Right:"
    },
    commonMarginLabel: {
        ja: "すべて同じ値にする",
        en: "Same Value"
    },
    guideExtensionLabel: {
        ja: "ガイドの伸張：",
        en: "Guide Extension:"
    },
    bleedGuideLabel: {
        ja: "裁ち落としのガイド：",
        en: "Bleed Guide:"
    },
    allBoardsLabel: {
        ja: "すべてのアートボードに適用",
        en: "Apply to All Artboards"
    },
    cellRectLabel: {
        ja: "セルを長方形化",
        en: "Create Cell Rectangles"
    },
    drawGuidesLabel: {
        ja: "ガイドを引く",
        en: "Draw Guides"
    },
    clearGuidesLabel: {
        ja: "grid_guidesレイヤーをクリア",
        en: "Clear grid_guides Layer"
    },
    cancelLabel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    applyLabel: {
        ja: "適用",
        en: "Apply"
    },
    okLabel: {
        ja: "OK",
        en: "OK"
    },
    exportPresetLabel: {
        ja: "書き出し",
        en: "Export"
    }
};

// プリセット定義（drawGuides, drawBleedGuide追加） / Preset definitions (drawGuides, drawBleedGuide added)
    var presets = [{
            label: "十字 / Cross",
            x: 2,
            y: 2,
            ext: 0,
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            rowGutter: 0,
            colGutter: 0,
            drawCells: false,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "シングル / Single",
            x: 1,
            y: 1,
            ext: 50,
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
            rowGutter: 0,
            colGutter: 0,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "2行×2列 / 2 Rows × 2 Columns",
            x: 2,
            y: 2,
            ext: 20,
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            rowGutter: 50,
            colGutter: 50,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "1行×3列 / 1 Row × 3 Columns",
            x: 3,
            y: 1,
            ext: 0,
            top: 30,
            bottom: 30,
            left: 30,
            right: 30,
            rowGutter: 0,
            colGutter: 30,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "4行×4列 / 4 Rows × 4 Columns",
            x: 4,
            y: 4,
            ext: 0,
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            rowGutter: 20,
            colGutter: 20,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "2行×3列 / 2 Rows × 3 Columns",
            x: 3,
            y: 2,
            ext: 0,
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
            rowGutter: 20,
            colGutter: 20,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "3行×3列 / 3 Rows × 3 Columns",
            x: 3,
            y: 3,
            ext: 0,
            top: 0,
            bottom: 0,
            left: 200,
            right: 0,
            rowGutter: 0,
            colGutter: 0,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "sp / sp",
            x: 1,
            y: 1,
            ext: 0,
            top: 220,
            bottom: 220,
            left: 0,
            right: 0,
            rowGutter: 0,
            colGutter: 0,
            drawCells: true,
            drawGuides: true,
            drawBleedGuide: false
        },
        {
            label: "長方形のみ / just rectangle",
            x: 1,
            y: 1,
            ext: 10,
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            rowGutter: 0,
            colGutter: 0,
            drawCells: true,
            drawGuides: false,
            drawBleedGuide: false
        }
    ];

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。\nPlease open a document.");
        return;
    }

    var doc = app.activeDocument;
    var rulerUnit = app.preferences.getIntegerPreference("rulerType");
    var unitLabel = "pt";
    var unitFactor = 1.0;

    // 単位設定
    if (rulerUnit === 0) {
        unitLabel = "inch";
        unitFactor = 72.0;
    } else if (rulerUnit === 1) {
        unitLabel = "mm";
        unitFactor = 72.0 / 25.4;
    } else if (rulerUnit === 2) {
        unitLabel = "pt";
        unitFactor = 1.0;
    } else if (rulerUnit === 3) {
        unitLabel = "pica";
        unitFactor = 12.0;
    } else if (rulerUnit === 4) {
        unitLabel = "cm";
        unitFactor = 72.0 / 2.54;
    } else if (rulerUnit === 5) {
        unitLabel = "Q";
        unitFactor = 72.0 / 25.4 * 0.25;
    } else if (rulerUnit === 6) {
        unitLabel = "px";
        unitFactor = 1.0;
    }
    // ダイアログ作成 / Create dialog
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    // grid_guidesレイヤークリアチェックボックス / Clear grid_guides layer checkbox
    var clearGuidesCheckbox = dlg.add("checkbox", undefined, LABELS.clearGuidesLabel[lang]);
    clearGuidesCheckbox.value = true;

    // プリセット選択＋書き出しボタングループ / Preset selection and export group
    var presetGroup = dlg.add("group");
    presetGroup.orientation = "row";
    presetGroup.alignChildren = "center";
    presetGroup.margins = [0, 5, 0, 10];

    presetGroup.add("statictext", undefined, LABELS.presetLabel[lang]);
    var presetDropdown = presetGroup.add("dropdownlist", undefined, []);
    presetDropdown.selection = 0;
    var btnExportPreset = presetGroup.add("button", undefined, LABELS.exportPresetLabel[lang]);

    btnExportPreset.onClick = function() {
        var saveFile = File.saveDialog("プリセットを書き出す場所と名前を指定してください / Choose where to save the preset", "*.txt");
        if (!saveFile) {
            return;
        }

        // 拡張子がない場合は.txtをつける / Add .txt extension if missing
        if (saveFile.name.indexOf(".") === -1) {
            saveFile = new File(saveFile.fsName + ".txt");
        }

        // ★ファイル名から.txtを正しく除去！ / Remove .txt extension from file name
        var fileName = saveFile.name.replace(/\.txt$/i, "");

        var currentPreset = {
            x: parseInt(inputXText.text, 10),
            y: parseInt(inputYText.text, 10),
            ext: parseFloat(inputExt.text),
            top: parseFloat(inputTop.text),
            bottom: parseFloat(inputBottom.text),
            left: parseFloat(inputLeft.text),
            right: parseFloat(inputRight.text),
            rowGutter: parseFloat(inputRowGutter.text),
            colGutter: parseFloat(inputColGutter.text),
            drawCells: cellRectCheckbox.value,
            drawGuides: drawGuidesCheckbox.value,
            drawBleedGuide: bleedGuideCheckbox.value
        };

        var presetString = '{ label: "' + fileName + '", ' +
            'x: ' + currentPreset.x + ', ' +
            'y: ' + currentPreset.y + ', ' +
            'ext: ' + currentPreset.ext + ', ' +
            'top: ' + currentPreset.top + ', ' +
            'bottom: ' + currentPreset.bottom + ', ' +
            'left: ' + currentPreset.left + ', ' +
            'right: ' + currentPreset.right + ', ' +
            'rowGutter: ' + currentPreset.rowGutter + ', ' +
            'colGutter: ' + currentPreset.colGutter + ', ' +
            'drawCells: ' + currentPreset.drawCells + ', ' +
            'drawGuides: ' + currentPreset.drawGuides + ', ' +
            'drawBleedGuide: ' + currentPreset.drawBleedGuide +
            ' }';

        if (saveFile.open("w")) {
            saveFile.write(presetString);
            saveFile.close();
            alert("プリセットを書き出しました！ / Preset exported!");
        } else {
            alert("ファイルを書き込めませんでした。 / Failed to write the file.");
        }
    };

    // グリッド設定グループ / Grid settings group
    var gridGroup = dlg.add("group");
    gridGroup.orientation = "row";
    gridGroup.alignChildren = "top";
    gridGroup.spacing = 20;

    // 行設定パネル / Row settings panel
    var rowBlock = gridGroup.add("panel", undefined, LABELS.rowTitle[lang]);
    rowBlock.orientation = "column";
    rowBlock.alignChildren = "left";
    rowBlock.margins = [15, 20, 15, 15];

    var inputY = rowBlock.add("group");
    inputY.add("statictext", undefined, LABELS.rowsLabel[lang]);
    var inputYText = inputY.add("edittext", undefined, "2");
    inputYText.characters = 3;

    var rowGutterGroup = rowBlock.add("group");
    rowGutterGroup.add("statictext", undefined, LABELS.rowGutterLabel[lang] + "：");
    var inputRowGutter = rowGutterGroup.add("edittext", undefined, "0");
    inputRowGutter.characters = 4;
    rowGutterGroup.add("statictext", undefined, unitLabel);

    // 列設定パネル / Column settings panel
    var colBlock = gridGroup.add("panel", undefined, LABELS.columnTitle[lang]);
    colBlock.orientation = "column";
    colBlock.alignChildren = "left";
    colBlock.margins = [15, 20, 15, 15];

    var inputX = colBlock.add("group");
    inputX.add("statictext", undefined, LABELS.columnsLabel[lang]);
    var inputXText = inputX.add("edittext", undefined, "2");
    inputXText.characters = 3;

    var colGutterGroup = colBlock.add("group");
    colGutterGroup.add("statictext", undefined, LABELS.colGutterLabel[lang] + "：");
    var inputColGutter = colGutterGroup.add("edittext", undefined, "0");
    inputColGutter.characters = 4;
    colGutterGroup.add("statictext", undefined, unitLabel);

    // マージン全体パネル / Margin panel
    var marginPanel = dlg.add("panel", undefined, LABELS.marginTitle[lang] + " (" + unitLabel + ")");
    marginPanel.orientation = "column";
    marginPanel.alignChildren = "left";
    var marginLabelWidth = (lang === "ja") ? 30 : 70; // unify Top/Bottom label width and right-align
    marginPanel.margins = [10, 15, 10, 15];

    // --- 上段グループ（左／上下／右） / Upper group (left/up-down/right) ---
    var upperGroup = marginPanel.add("group");
    upperGroup.orientation = "row";
    upperGroup.alignChildren = "top";

    // --- 左だけグループ / Left only group ---
    var leftGroup = upperGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = "center";
    leftGroup.margins = [0, 12, 0, 10];

    leftGroup.add("statictext", undefined, LABELS.leftLabel[lang]);
    var inputLeft = leftGroup.add("edittext", undefined, "0");
    inputLeft.characters = 5;

    // --- 上下グループ / Up-down group ---
    var topBottomGroup = upperGroup.add("group");
    topBottomGroup.orientation = "column";
    topBottomGroup.alignChildren = "left";
    topBottomGroup.margins = [5, 0, 5, 0];

    var topGroup = topBottomGroup.add("group");
    topGroup.orientation = "row";
    var lblTop = topGroup.add("statictext", undefined, LABELS.topLabel[lang]);
    lblTop.justification = "right";
    lblTop.minimumSize.width = marginLabelWidth;
    lblTop.maximumSize.width = marginLabelWidth;
    var inputTop = topGroup.add("edittext", undefined, "0");
    inputTop.characters = 5;

    var bottomGroup = topBottomGroup.add("group");
    bottomGroup.orientation = "row";
    var lblBottom = bottomGroup.add("statictext", undefined, LABELS.bottomLabel[lang]);
    lblBottom.justification = "right";
    lblBottom.minimumSize.width = marginLabelWidth;
    lblBottom.maximumSize.width = marginLabelWidth;
    var inputBottom = bottomGroup.add("edittext", undefined, "0");
    inputBottom.characters = 5;

    // --- 右だけグループ / Right only group ---
    var rightGroup = upperGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = "center";
    rightGroup.margins = [0, 12, 0, 10];

    rightGroup.add("statictext", undefined, LABELS.rightLabel[lang]);
    var inputRight = rightGroup.add("edittext", undefined, "0");
    inputRight.characters = 5;

    // --- 共通マージングループ（下段に追加） / Common margin group (bottom) ---
    var commonGroup = marginPanel.add("group");
    commonGroup.orientation = "row";
    commonGroup.alignChildren = "center";
    commonGroup.margins = [0, 10, 0, 0];

    var commonMarginCheckbox = commonGroup.add("checkbox", undefined, LABELS.commonMarginLabel[lang]);
    var commonMarginInput = commonGroup.add("edittext", undefined, "0");
    commonMarginInput.characters = 5;

    // オプション設定グループ（ガイドの伸張・裁ち落としガイド・セル長方形化・ガイドを引く）/ Options group (guide extension, bleed guide, cell rectangle, draw guides)
    var optGroup = dlg.add("group");
    optGroup.orientation = "column";
    optGroup.alignChildren = "left";

    // ガイドの伸張設定 / Guide extension
    var extGroup = optGroup.add("group");
    extGroup.margins = [0, 0, 0, 10];
    extGroup.add("statictext", undefined, LABELS.guideExtensionLabel[lang]);
    var inputExt = extGroup.add("edittext", undefined, "10");
    inputExt.characters = 5;
    extGroup.add("statictext", undefined, unitLabel);

    // 裁ち落としガイド設定 / Bleed guide
    var bleedGroup = optGroup.add("group");
    bleedGroup.margins = [0, 0, 0, 10];
    var bleedGuideCheckbox = bleedGroup.add("checkbox", undefined, LABELS.bleedGuideLabel[lang]);
    bleedGuideCheckbox.value = false;
    var inputBleed = bleedGroup.add("edittext", undefined, "3");
    inputBleed.characters = 4;
    bleedGroup.add("statictext", undefined, "(mm)");
    // --- 矢印キーで値を増減する関数 ---
    // Shiftキー押下時は10の倍数スナップ
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;

            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;

                if (keyboard.shiftKey) {
                    // Shiftキー押下時は10の倍数にスナップ
                    value = Math.floor(value / 10) * 10;
                    delta = 10;
                }

                value += isUp ? delta : -delta;
                if (value < 0) value = 0; // 必要なら下限チェック

                event.preventDefault();
                editText.text = value;
                // 「すべて同じ値にする」ON時は同期してからプレビュー / When "Same Value" is ON, sync first
                try {
                    if (commonMarginCheckbox && commonMarginCheckbox.value) {
                        syncCommonMargin();
                    }
                } catch (e) {}
                // 入力変更を即時プレビューに反映 / Refresh preview immediately
                try { drawGuides(true); } catch (e) {}
            }
        });
    }

    // 入力値変更で即時プレビュー / Live preview on any input change
    function attachLivePreview(editText) {
        editText.onChanging = function() {
            try { drawGuides(true); } catch (e) {}
        };
    }

    // --- 各数値editTextに矢印キー増減機能を追加 ---
    changeValueByArrowKey(inputXText);
    changeValueByArrowKey(inputYText);
    changeValueByArrowKey(inputExt);
    changeValueByArrowKey(inputTop);
    changeValueByArrowKey(inputBottom);
    changeValueByArrowKey(inputLeft);
    changeValueByArrowKey(inputRight);
    changeValueByArrowKey(inputRowGutter);
    changeValueByArrowKey(inputColGutter);
    changeValueByArrowKey(commonMarginInput);
    changeValueByArrowKey(inputBleed);

    // --- 入力中の変更もリアルタイム反映 / Attach onChanging for live preview ---
    attachLivePreview(inputXText);
    attachLivePreview(inputYText);
    attachLivePreview(inputExt);
    attachLivePreview(inputTop);
    attachLivePreview(inputBottom);
    attachLivePreview(inputLeft);
    attachLivePreview(inputRight);
    attachLivePreview(inputRowGutter);
    attachLivePreview(inputColGutter);
    attachLivePreview(commonMarginInput);
    attachLivePreview(inputBleed);

    var allBoardsCheckbox = optGroup.add("checkbox", undefined, LABELS.allBoardsLabel[lang]);

    // セル長方形化・ガイドを引くを横並び / Cell rectangle and draw guides (side by side)
    var cellGuideGroup = optGroup.add("group");
    cellGuideGroup.orientation = "row";
    cellGuideGroup.alignChildren = "left";

    var cellRectCheckbox = cellGuideGroup.add("checkbox", undefined, LABELS.cellRectLabel[lang]);
    var drawGuidesCheckbox = cellGuideGroup.add("checkbox", undefined, LABELS.drawGuidesLabel[lang]);
    drawGuidesCheckbox.value = true;

    // === ボタンエリア（レイアウト変更版）/ Button area (layout updated)
    var outerGroup = dlg.add("group");
    outerGroup.alignment = ["fill", "top"];
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.margins = [0, 10, 0, 0];
    outerGroup.spacing = 0;

    // 左グループ（キャンセルボタン）/ Left group (Cancel button)
    var buttonLeftGroup = outerGroup.add("group");
    buttonLeftGroup.orientation = "row";
    buttonLeftGroup.alignChildren = "left";
    var btnCancel = buttonLeftGroup.add("button", undefined, LABELS.cancelLabel[lang], {
        name: "cancel"
    });

    // スペーサー（横に伸びる空白）/ Spacer (horizontal stretch)
    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 70;
    spacer.maximumSize.height = 0;

    // 右グループ（適用・OKボタン）/ Right group (Apply/OK buttons)
    var rightGroup2 = outerGroup.add("group");
    rightGroup2.alignment = ["right", "center"];
    rightGroup2.orientation = "row";
    rightGroup2.alignChildren = "right";
    rightGroup2.spacing = 10;
    var btnOK = rightGroup2.add("button", undefined, LABELS.okLabel[lang], {
        name: "ok"
    });
    btnOK.alignment = ["right", "center"];

    // 表示用ラベルをローカライズ / Localize display label for dropdown
    function presetDisplayLabel(raw) {
        // 日本語UIのときは「 / 」以降を隠す / In Japanese UI, hide text after " / "
        if (lang === "ja") return String(raw).replace(/\s*\/.*$/, "");
        return raw;
    }
    // プリセットをドロップダウンに追加 / Add presets to dropdown
    for (var i = 0; i < presets.length; i++) {
        var display = presetDisplayLabel(presets[i].label);
        presetDropdown.add("item", display);
    }
    presetDropdown.selection = 0;

    // プリセット選択時、入力値を反映
    presetDropdown.onChange = function() {
        var p = presets[presetDropdown.selection.index];
        inputXText.text = p.x;
        inputYText.text = p.y;
        inputExt.text = p.ext;
        inputTop.text = p.top;
        inputBottom.text = p.bottom;
        inputLeft.text = p.left;
        inputRight.text = p.right;
        inputRowGutter.text = (p.rowGutter !== undefined) ? p.rowGutter : "0";
        inputColGutter.text = (p.colGutter !== undefined) ? p.colGutter : "0";
        cellRectCheckbox.value = (typeof p.drawCells !== "undefined") ? p.drawCells : false;
        drawGuidesCheckbox.value = (typeof p.drawGuides !== "undefined") ? p.drawGuides : true;
        bleedGuideCheckbox.value = (typeof p.drawBleedGuide !== "undefined") ? p.drawBleedGuide : false;
        updateGutterEnable();
        syncCommonMargin();
        drawGuides(true);
    };
    // 「すべて同じにする」同期処理 / Sync for "Same Value"
    function syncCommonMargin() {
        if (commonMarginCheckbox.value) {
            var val = commonMarginInput.text;
            inputTop.text = val;
            inputBottom.text = val;
            inputLeft.text = val;
            inputRight.text = val;
            inputTop.enabled = false;
            inputBottom.enabled = false;
            inputLeft.enabled = false;
            inputRight.enabled = false;
        } else {
            inputTop.enabled = true;
            inputBottom.enabled = true;
            inputLeft.enabled = true;
            inputRight.enabled = true;
        }
        drawGuides(true);
    }
    commonMarginInput.onChanging = syncCommonMargin;
    commonMarginCheckbox.onClick = syncCommonMargin;

    // ガター有効無効切り替え / Enable/disable gutter fields
    function updateGutterEnable() {
        var xVal = parseInt(inputXText.text, 10);
        var yVal = parseInt(inputYText.text, 10);
        inputRowGutter.enabled = (yVal > 1);
        inputColGutter.enabled = (xVal > 1);
    }
    // ← ガター有効切替時も即時プレビュー
    inputXText.onChanging = inputYText.onChanging = function() {
        updateGutterEnable();
        try { drawGuides(true); } catch (e) {}
    };

    // 「ガイドを引く」切り替え：伸張・裁ち落としをディム制御 + プレビュー
    drawGuidesCheckbox.onClick = function() {
        var enable = drawGuidesCheckbox.value;
        inputExt.enabled = enable;
        bleedGuideCheckbox.enabled = enable;
        inputBleed.enabled = enable;
        drawGuides(true);
    };

    // ★未定義だったトグルにもプレビュー反映
    bleedGuideCheckbox.onClick = function(){ try { drawGuides(true); } catch (e) {} };
    cellRectCheckbox.onClick  = function(){ try { drawGuides(true); } catch (e) {} };
    allBoardsCheckbox.onClick = function(){ try { drawGuides(true); } catch (e) {} };

    // OKボタン押下時 / OK button pressed
    btnOK.onClick = function() {
        updateGutterEnable();
        if (clearGuidesCheckbox.value) {
            clearGuidesLayer();
        }
        dlg.close(1);
    };

    // ガイド＆セル長方形＆裁ち落としガイドを描画 / Draw guides, cell rectangles, and bleed guides
    function drawGuides(isPreview) {
        if (isPreview) {
            removePreviewGuides();
        }

        var xDiv = parseInt(inputXText.text, 10);
        var yDiv = parseInt(inputYText.text, 10);
        var ext = parseFloat(inputExt.text) * unitFactor;
        var top = parseFloat(inputTop.text) * unitFactor;
        var bottom = parseFloat(inputBottom.text) * unitFactor;
        var left = parseFloat(inputLeft.text) * unitFactor;
        var right = parseFloat(inputRight.text) * unitFactor;
        var rowGutter = parseFloat(inputRowGutter.text) * unitFactor;
        var colGutter = parseFloat(inputColGutter.text) * unitFactor;
        var bleed = parseFloat(inputBleed.text) * (72.0 / 25.4); // mm → pt換算
        var allBoards = allBoardsCheckbox.value;
        var drawCells = cellRectCheckbox.value;
        var drawGuidesNow = drawGuidesCheckbox.value;
        var drawBleedGuide = bleedGuideCheckbox.value;

        if (isNaN(xDiv) || xDiv <= 0 || isNaN(yDiv) || yDiv <= 0) return;

        var gridLayerName = isPreview ? "_Preview_Guides" : "grid_guides";
        var gridLayer;
        try {
            gridLayer = doc.layers.getByName(gridLayerName);
        } catch (e) {
            gridLayer = doc.layers.add();
            gridLayer.name = gridLayerName;
        }
        gridLayer.locked = false;

        var cellLayer = gridLayer;
        if (!isPreview && drawCells) {
            try {
                cellLayer = doc.layers.getByName("cell-rectangle");
            } catch (e) {
                cellLayer = doc.layers.add();
                cellLayer.name = "cell-rectangle";
            }
            cellLayer.locked = false;
        }
        for (var b = 0; b < doc.artboards.length; b++) {
            if (!allBoards && b !== doc.artboards.getActiveArtboardIndex()) continue;

            var ab = doc.artboards[b];
            var rect = ab.artboardRect;
            var abLeft = rect[0],
                abTop = rect[1],
                abRight = rect[2],
                abBottom = rect[3];
            var baseLeft = abLeft + left;
            var baseRight = abRight - right;
            var baseTop = abTop - top;
            var baseBottom = abBottom + bottom;

            var usableWidth = baseRight - baseLeft;
            var usableHeight = baseTop - baseBottom;
            var totalColGutter = (xDiv - 1) * colGutter;
            var totalRowGutter = (yDiv - 1) * rowGutter;
            var cellWidth = (usableWidth - totalColGutter) / xDiv;
            var cellHeight = (usableHeight - totalRowGutter) / yDiv;

            var guideLeft = abLeft - ext;
            var guideRight = abRight + ext;
            var guideTop = abTop + ext;
            var guideBottom = abBottom - ext;

            if (drawGuidesNow) {
                // 「中心」用の水平・垂直ガイドを追加
                if (xDiv === 1 && yDiv === 1) {
                    var centerX = (baseLeft + baseRight) / 2;
                    var centerY = (baseTop + baseBottom) / 2;

                    var vCenter = doc.pathItems.add();
                    vCenter.setEntirePath([
                        [centerX, abTop + ext],
                        [centerX, abBottom - ext]
                    ]);
                    vCenter.stroked = false;
                    vCenter.filled = false;
                    vCenter.guides = true;
                    vCenter.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    var hCenter = doc.pathItems.add();
                    hCenter.setEntirePath([
                        [abLeft - ext, centerY],
                        [abRight + ext, centerY]
                    ]);
                    hCenter.stroked = false;
                    hCenter.filled = false;
                    hCenter.guides = true;
                    hCenter.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                } else {
                    // 通常ガイド描画（行・列）
                    var y = baseTop;
                    var line = doc.pathItems.add();
                    line.setEntirePath([
                        [guideLeft, y],
                        [guideRight, y]
                    ]);
                    line.stroked = false;
                    line.filled = false;
                    line.guides = true;
                    line.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    for (var j = 0; j < yDiv; j++) {
                        y -= cellHeight;
                        var line1 = doc.pathItems.add();
                        line1.setEntirePath([
                            [guideLeft, y],
                            [guideRight, y]
                        ]);
                        line1.stroked = false;
                        line1.filled = false;
                        line1.guides = true;
                        line1.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                        if (j < yDiv - 1) {
                            y -= rowGutter;
                            var line2 = doc.pathItems.add();
                            line2.setEntirePath([
                                [guideLeft, y],
                                [guideRight, y]
                            ]);
                            line2.stroked = false;
                            line2.filled = false;
                            line2.guides = true;
                            line2.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                        }
                    }

                    y = baseBottom;
                    var lineBottom = doc.pathItems.add();
                    lineBottom.setEntirePath([
                        [guideLeft, y],
                        [guideRight, y]
                    ]);
                    lineBottom.stroked = false;
                    lineBottom.filled = false;
                    lineBottom.guides = true;
                    lineBottom.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    var x = baseLeft;
                    var vline = doc.pathItems.add();
                    vline.setEntirePath([
                        [x, guideTop],
                        [x, guideBottom]
                    ]);
                    vline.stroked = false;
                    vline.filled = false;
                    vline.guides = true;
                    vline.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    for (var i2 = 0; i2 < xDiv; i2++) {
                        x += cellWidth;
                        var vline1 = doc.pathItems.add();
                        vline1.setEntirePath([
                            [x, guideTop],
                            [x, guideBottom]
                        ]);
                        vline1.stroked = false;
                        vline1.filled = false;
                        vline1.guides = true;
                        vline1.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                        if (i2 < xDiv - 1) {
                            x += colGutter;
                            var vline2 = doc.pathItems.add();
                            vline2.setEntirePath([
                                [x, guideTop],
                                [x, guideBottom]
                            ]);
                            vline2.stroked = false;
                            vline2.filled = false;
                            vline2.guides = true;
                            vline2.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                        }
                    }

                    x = baseRight;
                    var vlineRight = doc.pathItems.add();
                    vlineRight.setEntirePath([
                        [x, guideTop],
                        [x, guideBottom]
                    ]);
                    vlineRight.stroked = false;
                    vlineRight.filled = false;
                    vlineRight.guides = true;
                    vlineRight.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                }
            }

            if (drawGuidesNow && drawBleedGuide) {
                // ★裁ち落としガイド
                var bleedRect = doc.pathItems.rectangle(
                    abTop + bleed,
                    abLeft - bleed,
                    (abRight - abLeft) + bleed * 2,
                    (abTop - abBottom) + bleed * 2
                );
                bleedRect.stroked = false;
                bleedRect.filled = false;
                bleedRect.guides = true;
                bleedRect.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
            }

            if (drawCells && cellLayer) {
                var startX = baseLeft;
                var startY = baseTop;
                for (var row = 0; row < yDiv; row++) {
                    var cellY = startY - (cellHeight + rowGutter) * row;
                    for (var col = 0; col < xDiv; col++) {
                        var cellX = startX + (cellWidth + colGutter) * col;
                        var rect2 = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                        rect2.stroked = false;
                        rect2.filled = true;
                        rect2.fillColor = createBlackColor();
                        rect2.opacity = 15;
                    }
                }
            }
        }

        if (!isPreview) {
            gridLayer.locked = true;
        }

        if (isPreview) {
            app.redraw();
        }
    }
    // 黒色作成（CMYK／RGB対応）/ Create black color (CMYK/RGB)
    function createBlackColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        } else {
            var rgb = new RGBColor();
            rgb.red = 0;
            rgb.green = 0;
            rgb.blue = 0;
            return rgb;
        }
    }

    // プレビューガイド削除 / Remove preview guides
    function removePreviewGuides() {
        try {
            var previewLayer = doc.layers.getByName("_Preview_Guides");
            if (previewLayer) previewLayer.remove();
        } catch (e) {}
    }

    // grid_guidesレイヤーのガイドだけ削除 / Remove only guides from grid_guides layer
    function clearGuidesLayer() {
        var guidesLayer = null;
        for (var i3 = 0; i3 < doc.layers.length; i3++) {
            if (doc.layers[i3].name === "grid_guides") {
                guidesLayer = doc.layers[i3];
                break;
            }
        }
        if (guidesLayer) {
            if (guidesLayer.locked) {
                guidesLayer.locked = false;
            }
            for (var j = guidesLayer.pageItems.length - 1; j >= 0; j--) {
                var item = guidesLayer.pageItems[j];
                if (item.guides) {
                    item.remove();
                }
            }
        }
    }

    // ダイアログ初期プレビュー＆終了時処理 / Initial dialog preview & post-process
    updateGutterEnable();
    syncCommonMargin();
    drawGuides(true);

    if (dlg.show() === 1) {
        removePreviewGuides();
        if (clearGuidesCheckbox.value) {
            clearGuidesLayer();
        }
        drawGuides(false);
    } else {
        removePreviewGuides();
    }
}

main();