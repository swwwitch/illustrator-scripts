#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


/*

### スクリプト名：

GridTextLayout.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 複数選択したテキストフレームの背面に K15% の長方形を作成し、「_bg-rectangle」レイヤーに配置します。
- 行間・列間・領域サイズを調整しながらリアルタイムにプレビュー可能です。

### 主な機能：

- 選択テキストフレームを基にグリッドを自動算出
- 行間・列間（ガター）の調整
- 領域サイズのカスタマイズ
- プレビューで即時反映
- 不要な長方形の削除オプション

### 処理の流れ：

1. テキストフレームを選択
2. 行数・列数を自動判定
3. ダイアログでガターや領域を設定
4. プレビューで確認
5. OK時に確定処理を実行

### 更新履歴：

- v1.0 (20250804) : 初期バージョン

---

### Script Name：

GridTextLayout.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### Overview：

- Creates K15% rectangles behind selected text frames on a "_bg-rectangle" layer.
- Allows real-time preview while adjusting row/column gaps and region size.

### Main Features：

- Auto-detects grid layout from selected text frames
- Adjustable row and column gaps (gutters)
- Customizable region size
- Instant preview updates
- Option to delete background rectangles

### Process Flow：

1. Select text frames
2. Auto-detect rows and columns
3. Configure gutters and region size in dialog
4. Preview updates in real time
5. Confirm to finalize

### Update History：

- v1.0 (20250804) : Initial version. 

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "グリッド化 " + SCRIPT_VERSION,
        en: "Grid " + SCRIPT_VERSION
    },
    judgeTitle: {
        ja: "判定",
        en: "Detection"
    },
    rowJudge: {
        ja: "行",
        en: "Row"
    },
    colJudge: {
        ja: "列",
        en: "Column"
    },
    gutterTitle: {
        ja: "ガター",
        en: "Gutter"
    },
    regionTitle: {
        ja: "領域",
        en: "Region"
    },
    rowGap: {
        ja: "行間",
        en: "Row Gap"
    },
    colGap: {
        ja: "列間",
        en: "Column Gap"
    },
    width: {
        ja: "幅",
        en: "Width"
    },
    height: {
        ja: "高さ",
        en: "Height"
    },
    link: {
        ja: "連動",
        en: "Link"
    },
    optionTitle: {
        ja: "オプション",
        en: "Options"
    },
    deleteRect: {
        ja: "長方形を削除",
        en: "Delete Rectangles"
    },
    convertArea: {
        ja: "エリア内文字に変換",
        en: "Convert to Area Text"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
}

/* Illustrator用のJavaScript / JavaScript for Illustrator */
/* 単位コードとラベルのマップ / Unit code to label map */
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

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/* 選択テキストフレームから行数・列数を算出 / Calculate rows and columns from selected text frames */
function calculateRowAndCol(rects, toleranceY, toleranceX) {
    var allBounds = [Number.MAX_VALUE, -Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE];
    var rowTops = [];
    var colLefts = [];

    for (var i = 0; i < rects.length; i++) {
        var b = rects[i].visibleBounds;
        if (b[0] < allBounds[0]) allBounds[0] = b[0];
        if (b[1] > allBounds[1]) allBounds[1] = b[1];
        if (b[2] > allBounds[2]) allBounds[2] = b[2];
        if (b[3] < allBounds[3]) allBounds[3] = b[3];

        /* 行判定 / Row detection */
        var addedRow = false;
        for (var j = 0; j < rowTops.length; j++) {
            if (Math.abs(b[1] - rowTops[j]) < toleranceY) {
                addedRow = true;
                break;
            }
        }
        if (!addedRow) rowTops.push(b[1]);

        /* 列判定 / Column detection */
        var addedCol = false;
        for (var k = 0; k < colLefts.length; k++) {
            if (Math.abs(b[0] - colLefts[k]) < toleranceX) {
                addedCol = true;
                break;
            }
        }
        if (!addedCol) colLefts.push(b[0]);
    }

    var rowCount = rowTops.length;
    var colCount = colLefts.length;

    return {
        allBounds: allBounds,
        rowCount: rowCount,
        colCount: colCount
    };
}

function main() {
    try {
        var doc = app.activeDocument;
        if (!doc) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var sel = doc.selection;
        if (!sel || sel.length === 0) {
            alert("テキストを選択してください。");
            return;
        }

        var rects = [];
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "TextFrame") {
                rects.push(sel[i]);
            }
        }

        if (rects.length === 0) {
            alert("テキストフレームが選択されていません。");
            return;
        }

        /* 背景レイヤーを取得または作成 / Get or create background layer */
        var bgLayer;
        try {
            bgLayer = doc.layers.getByName("_bg-rectangle");
        } catch (e) {
            bgLayer = doc.layers.add();
            bgLayer.name = "_bg-rectangle";
            bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
        }

        /* 許容値設定 (グローバル定義へ移動) / Tolerance settings (moved to global definition) */
        // var toleranceY = 6;
        // var toleranceX = 6;
        var gridInfo = calculateRowAndCol(rects, toleranceY, toleranceX);
        var allBounds = gridInfo.allBounds;
        var rowCount = gridInfo.rowCount;
        var colCount = gridInfo.colCount;

        /* ダイアログボックスで余白を設定 / Dialog for grid margin settings */
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);

        /* --- ダイアログ位置・透過度調整コード / Dialog position and opacity adjustment --- */
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function() {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }
        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }
        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);
        /* --- ここまで / End --- */

        dlg.orientation = "column";
        dlg.alignChildren = "left";

        /* 判定設定パネル / Detection panel */
        var judgePanel = dlg.add("panel", undefined, LABELS.judgeTitle[lang]);
        judgePanel.orientation = "row";
        judgePanel.alignChildren = "top";
        judgePanel.margins = [15, 20, 15, 10];

        var judgeGroup = judgePanel.add("group");
        judgeGroup.orientation = "row";
        judgeGroup.alignChildren = ["left", "center"]; // 左寄せ・天地中央 / Left aligned, vertically centered

        var tolLabelY = judgeGroup.add("statictext", undefined, LABELS.rowJudge[lang]);
        var tolInputY = judgeGroup.add("edittext", undefined, String(toleranceY));
        tolInputY.characters = 3;

        var tolLabelX = judgeGroup.add("statictext", undefined, LABELS.colJudge[lang]);
        var tolInputX = judgeGroup.add("edittext", undefined, String(toleranceX));
        tolInputX.characters = 3;

        /* 入力変更で tolerance 値を更新し、↑↓キーにも対応 / Update tolerance value on input change, support up/down keys */
        handleInput(tolInputY, {
            previewFunc: function() {
                var val = Number(tolInputY.text);
                if (!isNaN(val) && val >= 0) {
                    toleranceY = val;
                    var gridInfo = calculateRowAndCol(rects, toleranceY, toleranceX);
                    allBounds = gridInfo.allBounds;
                    rowCount = gridInfo.rowCount;
                    colCount = gridInfo.colCount;
                    updatePreview();
                }
            }
        });
        handleInput(tolInputX, {
            previewFunc: function() {
                var val = Number(tolInputX.text);
                if (!isNaN(val) && val >= 0) {
                    toleranceX = val;
                    var gridInfo = calculateRowAndCol(rects, toleranceY, toleranceX);
                    allBounds = gridInfo.allBounds;
                    rowCount = gridInfo.rowCount;
                    colCount = gridInfo.colCount;
                    updatePreview();
                }
            }
        });

        /* ガター設定パネル / Gutter panel */
        var currentUnitLabel = getCurrentUnitLabel();
        var gutterPanel = dlg.add("panel", undefined, LABELS.gutterTitle[lang] + "（" + currentUnitLabel + "）");
        gutterPanel.orientation = "row";
        gutterPanel.alignChildren = "top";
        gutterPanel.margins = [15, 20, 15, 5];

        /* 領域設定パネル / Region settings panel */
        var regionPanel = dlg.add("panel", undefined, LABELS.regionTitle[lang] + "（" + currentUnitLabel + "）");
        regionPanel.orientation = "row";
        regionPanel.alignChildren = "top";
        regionPanel.margins = [15, 20, 15, 10];

        var labelWidth = 30; /* ラベルの共通幅 / Common label width */
        var sizeGroup = regionPanel.add("group");
        sizeGroup.orientation = "row";
        sizeGroup.alignChildren = "top";

        var widthGroup = sizeGroup.add("group");
        var widthLabel = widthGroup.add("statictext", undefined, LABELS.width[lang]);
        widthLabel.justify = "right";
        widthLabel.minimumSize.width = labelWidth;
        var widthInput = widthGroup.add("edittext", undefined, String(Math.round(allBounds[2] - allBounds[0])));
        widthInput.characters = 5;

        var heightGroup = sizeGroup.add("group");
        var heightLabel = heightGroup.add("statictext", undefined, LABELS.height[lang]);
        heightLabel.justify = "right";
        heightLabel.minimumSize.width = labelWidth;
        var heightInput = heightGroup.add("edittext", undefined, String(Math.round(allBounds[1] - allBounds[3])));
        heightInput.characters = 5;

        /* オプション設定パネル / Options panel */
        var optionPanel = dlg.add("panel", undefined, LABELS.optionTitle[lang]);
        optionPanel.orientation = "column";
        optionPanel.alignChildren = "left";
        optionPanel.margins = [15, 20, 15, 10];

        var deleteRectCheckbox = optionPanel.add("checkbox", undefined, LABELS.deleteRect[lang]);
        deleteRectCheckbox.value = false; // デフォルトOFF / Default OFF

        var convertToAreaTextCheckbox = optionPanel.add("checkbox", undefined, LABELS.convertArea[lang]);
        convertToAreaTextCheckbox.value = false;
        convertToAreaTextCheckbox.enabled = false; // ディム表示 / Dimmed

        /* 左カラム / Left column */
        var gutterLeft = gutterPanel.add("group");
        gutterLeft.orientation = "column";
        gutterLeft.alignChildren = "left";

        var rowGroup = gutterLeft.add("group");
        rowGroup.add("statictext", undefined, LABELS.rowGap[lang]);
        var rowInput = rowGroup.add("edittext", undefined, "1");
        rowInput.characters = 3;

        var colGroup = gutterLeft.add("group");
        colGroup.add("statictext", undefined, LABELS.colGap[lang]);
        var colInput = colGroup.add("edittext", undefined, "1");
        colInput.characters = 3;

        handleInput(rowInput, {
            linkInput: colInput,
            linkCheckbox: linkCheckbox,
            previewFunc: updatePreview
        });
        handleInput(colInput, {
            previewFunc: updatePreview
        });
        handleInput(widthInput, {
            previewFunc: updatePreview
        });
        handleInput(heightInput, {
            previewFunc: updatePreview
        });

        /* フォーカスを「行間」入力欄に設定 / Set focus to Row Gap input when dialog opens */
        dlg.onShow = function() {
            try {
                rowInput.active = true; // 行間にフォーカス / Focus on Row Gap
            } catch (e) {}
        };

        /* 右カラム / Right column */
        var gutterRight = gutterPanel.add("group");
        gutterRight.orientation = "column";
        gutterRight.alignChildren = ["fill", "fill"];

        var spacerTop = gutterRight.add("statictext", undefined, "");
        spacerTop.minimumSize.height = 10;

        var linkCheckbox = gutterRight.add("checkbox", undefined, LABELS.link[lang]);
        linkCheckbox.value = true; // デフォルトON / Default ON
        colInput.enabled = false; // 初期状態で列間をディム表示 / Initially dim Column Gap
        colInput.text = rowInput.text; // 行間と同期 / Sync with Row Gap

        var spacerBottom = gutterRight.add("statictext", undefined, "");
        spacerBottom.minimumSize.height = 10;

        /* 「連動」チェックで列間をディム表示 / Disable col gap input when 'Link' checked */
        linkCheckbox.onClick = function() {
            colInput.enabled = !linkCheckbox.value;
            if (linkCheckbox.value) {
                colInput.text = rowInput.text; // 値を同期 / Sync value
            }
        };
        rowInput.onChange = function() {
            if (linkCheckbox.value) {
                colInput.text = rowInput.text;
            }
        };

        /* --- LIVE PREVIEW INSERTION START --- */
        /* プレビュー用長方形リスト / Preview rectangles list */
        var previewRects = [];

        /* プレビューをクリア / Clear preview rectangles */
        function clearPreview() {
            for (var i = 0; i < previewRects.length; i++) {
                try {
                    previewRects[i].remove();
                } catch (e) {}
            }
            previewRects = [];
        }

        /* グローバル変数として宣言 / Declare as global variables */
        var cellW, cellH;
        var rowGapVal, colGapVal;

        function updatePreview() {
            clearPreview();

            rowGapVal = Number(rowInput.text);
            colGapVal = Number(colInput.text);
            if (isNaN(rowGapVal)) rowGapVal = 0;
            if (isNaN(colGapVal)) colGapVal = 0;

            var totalW = Number(widthInput.text);
            var totalH = Number(heightInput.text);
            if (isNaN(totalW) || totalW <= 0) totalW = allBounds[2] - allBounds[0];
            if (isNaN(totalH) || totalH <= 0) totalH = allBounds[1] - allBounds[3];

            /* ここでグローバル変数を更新 / Update global variables here */
            cellW = (totalW - (colCount - 1) * colGapVal) / colCount;
            cellH = (totalH - (rowCount - 1) * rowGapVal) / rowCount;

            for (var r = 0; r < rowCount; r++) {
                for (var c = 0; c < colCount; c++) {
                    var left = allBounds[0] + c * (cellW + colGapVal);
                    var top = allBounds[1] - r * (cellH + rowGapVal);

                    var rect = bgLayer.pathItems.rectangle(top, left, cellW, cellH);
                    rect.stroked = false;
                    rect.filled = true;
                    rect.fillColor = makeGrayColor(15);
                    rect.zOrder(ZOrderMethod.SENDTOBACK);
                    previewRects.push(rect);
                }
            }

            centerTextInCells(sel, rowCount, colCount, allBounds, rowGapVal, colGapVal, cellW, cellH);
            app.redraw();
        }

        /* ボタングループ作成とOK/キャンセルボタン / Button group with OK/Cancel */
        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "center";

        var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang], {
            name: "cancel"
        });
        cancelBtn.onClick = function() {
            clearPreview();
            dlg.close();
        };

        var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });
        okBtn.onClick = function() {
            /* プレビュー結果を確定 / Commit preview result */
            updatePreview();
            /*
            if (convertToAreaTextCheckbox.value) {
                // エリア内文字化処理は現在無効化されています / Area text conversion is currently disabled
            }
            */
            if (deleteRectCheckbox.value) {
                clearPreview();
                try {
                    var layerToDelete = doc.layers.getByName("_bg-rectangle");
                    layerToDelete.remove();
                } catch (e) {}
            }
            dlg.close();
        };

        /* 入力変更でプレビュー更新 / Update preview on input change */
        rowInput.onChanging = function() {
            if (linkCheckbox.value) colInput.text = rowInput.text;
            updatePreview();
        };
        colInput.onChanging = function() {
            updatePreview();
        };
        widthInput.onChanging = updatePreview;
        heightInput.onChanging = updatePreview;

        /* 初期プレビュー / Initial preview */
        updatePreview();
        /* --- LIVE PREVIEW INSERTION END --- */

        if (dlg.show() != 1) {
            return; // キャンセル時は終了 / Exit on cancel
        }
        // OK時はプレビューで描画・中央配置済みなので何もしない / No action needed on OK, already previewed and centered
    } catch (err) {
        alert("エラーが発生しました: " + err);
    }
}

/* K指定のグレースケールカラーを作成 / Create grayscale color with specified K value */
function makeGrayColor(k) {
    var gray = new GrayColor();
    gray.gray = k;
    return gray;
}

/* 入力欄をチェックボックスやリンク対象と連動させ、矢印キー操作にも対応 / Unified input handler */
function handleInput(editText, options) {
    var linkInput = options && options.linkInput ? options.linkInput : null;
    var linkCheckbox = options && options.linkCheckbox ? options.linkCheckbox : null;
    var previewFunc = options && options.previewFunc ? options.previewFunc : null;

    // onChangingでプレビュー更新
    editText.onChanging = function() {
        if (linkCheckbox && linkCheckbox.value && linkInput) {
            linkInput.text = editText.text;
        }
        if (previewFunc) previewFunc();
    };

    // ↑↓キー対応
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
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
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
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

        value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
        editText.text = value;

        if (linkCheckbox && linkCheckbox.value && linkInput) {
            linkInput.text = editText.text;
        }

        if (previewFunc) {
            app.redraw();
            previewFunc();
        }
    });
}

/* 入力値をリンク対象に反映してプレビューも更新 / Update linked input and preview */
function updateLinkedInputAndPreview(sourceInput, targetInput, previewFunc) {
    targetInput.text = sourceInput.text;
    if (previewFunc) previewFunc();
}

/* セル内でテキストを中央に配置 / Center text within its assigned cell */
function centerTextInCells(sel, rowCount, colCount, allBounds, rowGapVal, colGapVal, cellW, cellH) {
    for (var t = 0; t < sel.length; t++) {
        if (sel[t].typename !== "TextFrame") continue;
        var tfBounds = sel[t].visibleBounds;
        var tfCenterX = (tfBounds[0] + tfBounds[2]) / 2;
        var tfCenterY = (tfBounds[1] + tfBounds[3]) / 2;

        for (var r = 0; r < rowCount; r++) {
            for (var c = 0; c < colCount; c++) {
                var left = allBounds[0] + c * (cellW + colGapVal);
                var top = allBounds[1] - r * (cellH + rowGapVal);

                if (tfCenterX >= left && tfCenterX <= left + cellW &&
                    tfCenterY <= top && tfCenterY >= top - cellH) {

                    var tfW = tfBounds[2] - tfBounds[0];
                    var tfH = tfBounds[1] - tfBounds[3];
                    var newLeft = left + (cellW - tfW) / 2;
                    var newTop = top - (cellH - tfH) / 2;

                    var dx = newLeft - tfBounds[0];
                    var dy = newTop - tfBounds[1];
                    sel[t].translate(dx, dy);
                    break;
                }
            }
        }
    }
}

/* toleranceY, toleranceX グローバル宣言 / Global declaration */
var toleranceY = 6;
var toleranceX = 6;

main();