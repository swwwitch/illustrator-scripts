#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

Slice2Artboards.jsx

### 概要

- 選択した画像やオブジェクトを指定した行数・列数に分割し、各ピースを矩形マスクでクリッピングします。
- さらに各ピースをアートボードとして変換・生成できます。
- 印刷面付け、パズル風レイアウト、複数アートボード化に便利です。

### 主な機能

- グリッド分割（行数・列数指定）
- オフセットによるサイズ調整
- アスペクト比選択（A4、スクエア、US Letter、16:9、8:9、カスタム）
- アートボードへの自動変換、名前設定、ゼロ埋め
- マージン設定

### 処理の流れ

1. ダイアログで行数・列数、形状（アスペクト比）などを設定
2. OK実行時に分割用マスクを生成
3. 必要に応じてアートボードを追加・リネーム
4. 元画像の削除（オプション）

### 処理の流れ

このスクリプトを実行後、次の流れを想定しています。

1. Illustratorの標準の機能でアートボードを再配置
2. ResizeClipMaskスクリプトでマスクパスの大きさを調整
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx


### 更新履歴

- v1.0 (20250710): 初期バージョン
- v1.1 (20250710): アートボード変換とオプション追加
- v1.2 (20250710): 形状バリエーション追加、カスタム設定対応
- v1.3 (20250710): 微調整
- v1.4 (20250710): アスペクト比の候補を調整

### Script Name:

Slice2Artboards.jsx

### Overview

- Splits selected image or objects into grid pieces using specified rows and columns, masking each piece with a rectangle.
- Each piece can be converted into an artboard automatically.
- Useful for imposition, puzzle layouts, and creating multiple artboards.

### Features

- Grid splitting with row/column settings
- Offset-based size adjustments
- Aspect ratio presets (A4, Square, US Letter, 16:9, 8:9, Custom)
- Automatic conversion to artboards, naming, zero padding
- Margin setting

### Workflow

1. Configure grid, aspect ratio, and options in the dialog
2. Generate grid masks on execution
3. Optionally add and rename artboards
4. Delete original artwork (optional)

### Workflow
After running this script, the expected workflow is:

1. Rearrange artboards using Illustrator's standard features
2. Adjust mask path size using the ResizeClipMask script
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx

### Change Log

- v1.0 (20250710): Initial version
- v1.1 (20250710): Added artboard conversion and options
- v1.2 (20250710): Added shape variations and custom settings
- v1.3 (20250710): Minor adjustments
- v1.4 (20250710): Adjusted aspect ratio options

*/

var SCRIPT_VERSION = "v1.3";

(function(global) {
    /**
     * 現在のロケールから日本語か英語かを判定 / Determine language by locale
     * @return {string} 'ja' or 'en'
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    // -------------------------------
    // 日英ラベル定義 / Define labels (JA/EN)
    // -------------------------------
    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle: {
            ja: "分割してアートボード化",
            en: "Slice and Create Artboards"
        },
        shapePanel: {
            ja: "アスペクト比",
            en: "Aspect Ratio"
        },
        shapeA4: {
            ja: "A4 (210 x 297)",
            en: "A4 (210 x 297)"
        },
        shapeSquare: {
            ja: "スクエア",
            en: "Square"
        },
        shapeLetter: {
            ja: "US Letter",
            en: "US Letter"
        },
        shapeLegal: {
            ja: "US Legal (8.5 x 14)",
            en: "US Legal (8.5 x 14)"
        },
        shapeTabloid: {
            ja: "Tabloid (11 x 17)",
            en: "Tabloid (11 x 17)"
        },
        shape169: {
            ja: "16:9",
            en: "16:9"
        },
        shape89: {
            ja: "8:9",
            en: "8:9"
        },
        columns: {
            ja: "列数",
            en: "Columns"
        },
        rows: {
            ja: "行数",
            en: "Rows"
        },
        offsetLabel: {
            ja: "オフセット",
            en: "Offset"
        },
        okBtn: {
            ja: "実行",
            en: "Run"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        convertArtboard: {
            ja: "アートボードに変換",
            en: "Convert to Artboards"
        },
        options: {
            ja: "オプション",
            en: "Options"
        },
        artboardName: {
            ja: "アートボード名：",
            en: "Artboard Name:"
        },
        startNumber: {
            ja: "連番の初期値：",
            en: "Starting Number:"
        },
        zeroPad: {
            ja: "ゼロ埋め",
            en: "Zero Padding"
        },
        margin: {
            ja: "マージン：",
            en: "Margin:"
        }
    };

    // グローバルへ公開
    global.getCurrentLang = getCurrentLang;
    global.LABELS = LABELS;
})(this);

// マジックナンバーの定数化 / Magic number constants
var DEFAULT_COLUMN_COUNT = 5;
var DEFAULT_ROW_COUNT = 5;
var DEFAULT_OFFSET = -20;
var DEFAULT_MARGIN = 0;
var DEFAULT_START_NUMBER = 1;

/*
定数・ユーティリティ群 / Constants and utility functions
*/

/**
 * 汎用例外ラッパー / General try-catch wrapper
 * @param {Function} fn
 * @param {string} errMsg
 */
function safeExecute(fn, errMsg) {
    try {
        fn();
    } catch (e) {
        alert((errMsg || "エラーが発生しました") + ":\n" + e.message);
    }
}

/**
 * 単位コードとラベルのマップ / Map from unit code to label
 */
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

/**
 * 現在の単位ラベルを取得 / Get current document unit label
 * @return {string}
 */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/**
 * 列・行数入力検証 / Validate grid input
 * @param {number} columnCount
 * @param {number} rowCount
 * @return {boolean}
 */
function validateGridInput(columnCount, rowCount) {
    return (
        (columnCount >= 1 && rowCount >= 1) ||
        (columnCount == 0 && rowCount >= 1) ||
        (columnCount >= 1 && rowCount == 0)
    );
}

/**
 * 矩形の座標をオフセット調整（左右・上下独立） / Offset rectangle bounds (independent X/Y)
 * @param {PathItem} pathItem
 * @param {number} offsetX
 * @param {number} offsetY
 * @return {PathItem}
 */
function offsetRectangle(pathItem, offsetX, offsetY) {
    var gb = pathItem.geometricBounds;
    var newWidth = (gb[2] - gb[0]) + offsetX * 2;
    var newHeight = (gb[1] - gb[3]) + offsetY * 2;
    pathItem.left = pathItem.left - offsetX;
    pathItem.top = pathItem.top + offsetY;
    pathItem.width = newWidth;
    pathItem.height = newHeight;
    return pathItem;
}

/*
ここから main() 以下のメイン処理 / Main processing below main()
*/

function main() {
    var lang = getCurrentLang();
    app.undoGroup = (lang === "ja" ? "アートボード分割と作成" : "Slice and Create Artboards");
    safeExecute(function() {
        var lang = getCurrentLang();

        /*
         ダイアログ作成 / Create dialog
        */
        var dialogTitle = LABELS.dialogTitle[lang] + " " + SCRIPT_VERSION;
        var Dialog = new Window('dialog', dialogTitle);
        Dialog.orientation = 'column';
        Dialog.alignment = 'right';

        /*
         メイングループ（2カラム） / Main group (2 columns)
        */
        var mainGroup = Dialog.add('group');
        mainGroup.orientation = 'row';
        mainGroup.alignChildren = 'top';

        /*
         左カラム / Left column
        */
        var inputPanel = mainGroup.add('group');
        inputPanel.orientation = 'column';
        inputPanel.alignChildren = 'left';

        /*
         形状パネル / Shape panel (A4/Square/Letter/16:9/8:9/Custom)
        */
        var shapePanel = inputPanel.add("panel", undefined, LABELS.shapePanel[lang]);
        shapePanel.orientation = "column";
        shapePanel.alignChildren = "left";
        shapePanel.margins = [15, 20, 15, 10];
        var shapeRadioA4 = shapePanel.add("radiobutton", undefined, LABELS.shapeA4[lang]);
        var shapeRadioSquare = shapePanel.add("radiobutton", undefined, LABELS.shapeSquare[lang]);
        if (lang !== "ja") {
            var shapeRadioLetter = shapePanel.add("radiobutton", undefined, LABELS.shapeLetter[lang]);
            var shapeRadioLegal = shapePanel.add("radiobutton", undefined, LABELS.shapeLegal[lang]);
            var shapeRadioTabloid = shapePanel.add("radiobutton", undefined, LABELS.shapeTabloid[lang]);
        }
        var shapeRadio169 = shapePanel.add("radiobutton", undefined, LABELS.shape169[lang]);
        var shapeRadio89 = shapePanel.add("radiobutton", undefined, LABELS.shape89[lang]);
        var shapeRadioCustom = shapePanel.add("radiobutton", undefined, (lang === "ja" ? "カスタム" : "Custom"));
        shapeRadioA4.value = true; // デフォルトA4 / Default is A4

        /*
         ピース設定グループ / Piece setting group
        */
        var pieceSettingGroup = inputPanel.add("group");
        pieceSettingGroup.orientation = "column";
        pieceSettingGroup.alignChildren = "left";
        pieceSettingGroup.margins = [10, 5, 10, 15];

        /*
         行数・列数入力欄 / Row and Column Inputs
        */
        var rowColGroup = pieceSettingGroup.add('group');
        rowColGroup.orientation = 'row';
        rowColGroup.alignChildren = 'left';
        // 列数 / Columns
        var colGroup = rowColGroup.add('group');
        colGroup.orientation = 'row';
        colGroup.add('statictext', undefined, LABELS.columns[lang]);
        var columnText = colGroup.add('edittext', undefined, String(DEFAULT_COLUMN_COUNT));
        columnText.characters = 3;
        // 行数 / Rows
        var rowGroup = rowColGroup.add('group');
        rowGroup.orientation = 'row';
        rowGroup.add('statictext', undefined, LABELS.rows[lang]);
        var rowText = rowGroup.add('edittext', undefined, String(DEFAULT_ROW_COUNT));
        rowText.characters = 3;

        /*
         オフセットグループ / Offset Group
        */
        var offsetGroup = inputPanel.add("group");
        offsetGroup.orientation = "row";
        offsetGroup.alignChildren = "left";
        offsetGroup.margins = [10, 0, 10, 10];
        var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
        offsetCheckbox.value = true;
        var offsetValueInput = offsetGroup.add("edittext", undefined, String(DEFAULT_OFFSET));
        offsetValueInput.characters = 4;
        var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
        offsetValueInput.enabled = offsetCheckbox.value;
        offsetUnitLabel.enabled = offsetCheckbox.value;
        offsetCheckbox.onClick = function() {
            offsetValueInput.enabled = offsetCheckbox.value;
            offsetUnitLabel.enabled = offsetCheckbox.value;
        };

        /*
         右カラム / Right column
        */
        var artboardColumnGroup = mainGroup.add('group');
        artboardColumnGroup.orientation = 'column';
        artboardColumnGroup.alignChildren = 'left';

        // artboardCheckbox を artboardColumnGroup 内に追加
        var artboardCheckbox = artboardColumnGroup.add('checkbox', undefined, LABELS.convertArtboard[lang]);
        artboardCheckbox.value = true;

        // artboardPanel を artboardColumnGroup 内に追加
        var artboardPanel = artboardColumnGroup.add("panel", undefined, LABELS.options[lang]);
        artboardPanel.orientation = "column";
        artboardPanel.alignChildren = "left";
        artboardPanel.margins = [15, 20, 15, 10];
        // --- アートボード名入力欄追加 / Add artboard name input ---
        var artboardNameGroup = artboardPanel.add("group");
        artboardNameGroup.orientation = "column";
        artboardNameGroup.add("statictext", undefined, LABELS.artboardName[lang]);
        var artboardNameInput = artboardNameGroup.add("edittext", undefined, "");
        artboardNameInput.characters = 10;
        // --- 連番の初期値入力欄追加 / Add starting number input ---
        var artboardNumberGroup = artboardPanel.add("group");
        artboardNumberGroup.orientation = "row";
        artboardNumberGroup.add("statictext", undefined, LABELS.startNumber[lang]);
        var artboardNumberInput = artboardNumberGroup.add("edittext", undefined, String(DEFAULT_START_NUMBER));
        artboardNumberInput.characters = 3;

        // --- ゼロ埋めチェックボックス追加 / Add zero padding checkbox ---
        var zeroPadCheckbox = artboardPanel.add('checkbox', undefined, LABELS.zeroPad[lang]);
        zeroPadCheckbox.value = true; // デフォルトON / Default ON

        // --- マージン入力欄追加 / Add margin input ---
        var artboardMarginGroup = artboardPanel.add("group");
        artboardMarginGroup.orientation = "row";
        artboardMarginGroup.add("statictext", undefined, LABELS.margin[lang]);
        var artboardMarginInput = artboardMarginGroup.add("edittext", undefined, String(DEFAULT_MARGIN));
        artboardMarginInput.characters = 5;
        // 単位ラベルを追加
        var artboardMarginUnitLabel = artboardMarginGroup.add("statictext", undefined, getCurrentUnitLabel());

        // --- マージン初期値をオフセット値に連動 / Link margin value to offset ---
        if (offsetCheckbox.value) {
            var offsetVal = parseFloat(offsetValueInput.text);
            if (!isNaN(offsetVal)) {
                artboardMarginInput.text = String(Math.abs(offsetVal));
            }
        }

        // --- 「アートボードに変換」チェックボックスで他オプション有効/無効切り替え / Enable/disable options by artboard checkbox ---
        artboardCheckbox.onClick = function() {
            var enabled = artboardCheckbox.value;
            artboardNameInput.enabled = enabled;
            artboardNumberInput.enabled = enabled;
            zeroPadCheckbox.enabled = enabled;
            artboardMarginInput.enabled = enabled;
        };

        /*
         OK・キャンセルボタン / OK & Cancel Buttons
        */
        var groupButtons = Dialog.add('group');
        groupButtons.orientation = 'row';
        groupButtons.alignment = "right";
        groupButtons.add('button', undefined, LABELS.cancel[lang]);
        var okBtn = groupButtons.add('button', undefined, LABELS.okBtn[lang], {
            name: "ok"
        });
        okBtn.active = true;

        /*
         選択画像やベクターアートワークがあれば初期値を自動設定 / Auto-set grid if artwork selected
        */
        function getSelectedArtworkItem() {
            if (app.documents.length > 0 && app.selection.length == 1) {
                var sel = app.selection[0];
                if (
                    sel.typename === "RasterItem" ||
                    sel.typename === "PlacedItem" ||
                    sel.typename === "SymbolItem" ||
                    sel.typename === "PathItem" ||
                    sel.typename === "GroupItem" ||
                    sel.typename === "CompoundPathItem"
                ) {
                    var gb = sel.geometricBounds;
                    var width = gb[2] - gb[0];
                    var height = gb[1] - gb[3];
                    if (height < 0) height = -height;
                    return {
                        width: width,
                        height: height
                    };
                }
            }
            return null;
        }
        /*
         行列数の自動計算（A4比率/Square/Letter/16:9/8:9） / Auto-calc grid by shape
        */
        var selectedImage = getSelectedArtworkItem();

        function updateGridByShape() {
            if (!selectedImage) return;
            // カスタムの場合は自動計算をスキップ
            if (shapeRadioCustom.value) {
                // カスタムの場合はユーザー入力優先（何もしない）
                return;
            }
            var aspect = selectedImage.width / selectedImage.height;
            var targetAspect = 1.0;
            if (shapeRadioA4.value) {
                targetAspect = 210 / 297;
            } else if (shapeRadioSquare.value) {
                targetAspect = 1.0;
            } else if (typeof shapeRadioLetter !== "undefined" && shapeRadioLetter.value) {
                targetAspect = 8.5 / 11;
            } else if (typeof shapeRadioLegal !== "undefined" && shapeRadioLegal.value) {
                targetAspect = 8.5 / 14;
            } else if (typeof shapeRadioTabloid !== "undefined" && shapeRadioTabloid.value) {
                targetAspect = 11 / 17;
            } else if (shapeRadio169.value) {
                targetAspect = 16 / 9;
            } else if (shapeRadio89.value) {
                targetAspect = 8 / 9;
            }
            if (aspect > targetAspect) {
                rowText.text = "1";
                columnText.text = String(Math.max(1, Math.round(aspect / targetAspect)));
            } else {
                columnText.text = "1";
                rowText.text = String(Math.max(1, Math.round(targetAspect / aspect)));
            }
        }
        if (selectedImage) {
            updateGridByShape();
        }
        shapeRadioA4.onClick = function() {
            updateGridByShape();
        };
        shapeRadioSquare.onClick = function() {
            updateGridByShape();
        };
        if (typeof shapeRadioLetter !== "undefined") {
            shapeRadioLetter.onClick = function() {
                updateGridByShape();
            };
        }
        if (typeof shapeRadioLegal !== "undefined") {
            shapeRadioLegal.onClick = function() {
                updateGridByShape();
            };
        }
        if (typeof shapeRadioTabloid !== "undefined") {
            shapeRadioTabloid.onClick = function() {
                updateGridByShape();
            };
        }
        shapeRadio169.onClick = function() {
            updateGridByShape();
        };
        shapeRadio89.onClick = function() {
            updateGridByShape();
        };
        shapeRadioCustom.onClick = function() {
            updateGridByShape();
        };
        offsetValueInput.onChanging = function() {};

        // 列数・行数の onChanging でカスタム選択 / Select custom when changing column/row
        columnText.onChanging = function() {
            shapeRadioCustom.value = true;
        };
        rowText.onChanging = function() {
            shapeRadioCustom.value = true;
        };

        // --- メイン処理（OK押下時） / Main logic on OK ---
        if (
            app.documents.length > 0 &&
            app.selection.length > 0 &&
            Dialog.show() == 1 &&
            (columnText.text || rowText.text)
        ) {
            var selectedObj;
            var origImageObj = null;
            var selectedObjectForPuzzle;
            var isTempRect = false;

            // シンボル化処理を共通関数で実行
            var symbolResult = convertToSymbol(app.selection);
            selectedObj = symbolResult.symbolItem;
            selectedObjectForPuzzle = symbolResult.maskRect ? symbolResult.maskRect : selectedObj;
            origImageObj = symbolResult.origImageObj ? symbolResult.origImageObj : null;
            isTempRect = symbolResult.isTempRect ? true : false;
            selectedImage = app.selection[0];

            // 列・行数取得 / Get row and column count
            var columnCount = Math.round(Number(columnText.text));
            var rowCount = Math.round(Number(rowText.text));

            // 入力値検証 / Validate input
            if (validateGridInput(columnCount, rowCount)) {
                selectedObjectForPuzzle.selected = false;

                // 列または行が0の場合は比率で自動算出 / Auto-calc if row/col is zero
                var aspectRatio;
                if (columnCount == 0) {
                    aspectRatio = Math.abs(
                        (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0]) /
                        (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1])
                    );
                    columnCount = Math.round(aspectRatio * rowCount);
                }
                if (rowCount == 0) {
                    aspectRatio = Math.abs(
                        (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1]) /
                        (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0])
                    );
                    rowCount = Math.round(aspectRatio * columnCount);
                }

                /*
                 マスク矩形のアスペクト比決定 / Mask aspect ratio by shape radio
                */
                var maskAspectRatio = 1.0;
                if (shapeRadioA4.value) {
                    maskAspectRatio = 210 / 297;
                } else if (shapeRadioSquare.value) {
                    maskAspectRatio = 1.0;
                } else if (typeof shapeRadioLetter !== "undefined" && shapeRadioLetter.value) {
                    maskAspectRatio = 8.5 / 11;
                } else if (typeof shapeRadioLegal !== "undefined" && shapeRadioLegal.value) {
                    maskAspectRatio = 8.5 / 14;
                } else if (typeof shapeRadioTabloid !== "undefined" && shapeRadioTabloid.value) {
                    maskAspectRatio = 11 / 17;
                } else if (typeof shapeRadio169 !== "undefined" && shapeRadio169.value) {
                    maskAspectRatio = 16 / 9;
                } else if (typeof shapeRadio89 !== "undefined" && shapeRadio89.value) {
                    maskAspectRatio = 8 / 9;
                }

                // オブジェクト左上座標（上端揃え） / Top-left coordinate (align to top)
                var originX = selectedObjectForPuzzle.geometricBounds[0];
                var originY = selectedObjectForPuzzle.geometricBounds[1];
                var objWidth = selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0];
                var objHeight = selectedObjectForPuzzle.geometricBounds[1] - selectedObjectForPuzzle.geometricBounds[3];
                if (objHeight < 0) objHeight = -objHeight;

                // --- カスタム選択時はアスペクト比ではなく、行列数で等分割 / If custom, ignore aspect ratio
                var pieceWidth, pieceHeight;
                var gridTotalWidth, gridTotalHeight;
                if (shapeRadioCustom.value) {
                    pieceWidth = objWidth / columnCount;
                    pieceHeight = objHeight / rowCount;
                    gridTotalWidth = pieceWidth * columnCount;
                    gridTotalHeight = pieceHeight * rowCount;
                } else {
                    // 現行の形状比率モード
                    var w0 = objWidth / columnCount;
                    var h0 = objHeight / rowCount;
                    if (w0 / h0 > maskAspectRatio) {
                        pieceHeight = h0;
                        pieceWidth = pieceHeight * maskAspectRatio;
                    } else {
                        pieceWidth = w0;
                        pieceHeight = pieceWidth / maskAspectRatio;
                    }
                    gridTotalWidth = pieceWidth * columnCount;
                    gridTotalHeight = pieceHeight * rowCount;
                }
                // originX, originY: 上端揃え（左上基準）
                originX = selectedObjectForPuzzle.geometricBounds[0];
                originY = selectedObjectForPuzzle.geometricBounds[1];

                /*
                 グリッド形状（矩形マスク）モード / Grid mask mode
                */
                var generatedPieces = [];
                for (var y = 0; y < rowCount; y++) {
                    for (var x = 0; x < columnCount; x++) {
                        var gridMask = app.activeDocument.pathItems.rectangle(
                            originY - y * pieceHeight,
                            originX + x * pieceWidth,
                            pieceWidth,
                            pieceHeight
                        );
                        // オフセット適用 / Apply offset if enabled
                        if (offsetCheckbox.value) {
                            var offsetVal = parseFloat(offsetValueInput.text);
                            gridMask = offsetRectangle(gridMask, offsetVal, offsetVal);
                        }
                        gridMask.closed = true;
                        gridMask.filled = false;
                        gridMask.stroked = false;
                        // 画像やシンボルならクリッピング / Clip if image or symbol
                        if (origImageObj && (origImageObj.typename === "PlacedItem" || origImageObj.typename === "SymbolItem")) {
                            var placedCopy = origImageObj.duplicate();
                            var group = app.activeDocument.groupItems.add();
                            placedCopy.moveToBeginning(group);
                            gridMask.moveToBeginning(group);
                            gridMask.clipping = true;
                            group.clipped = true;
                            generatedPieces.push(group);
                            group.selected = true;
                        } else {
                            generatedPieces.push(gridMask);
                            gridMask.selected = true;
                        }
                    }
                }

                // --- 追加: 生成したクリップグループの重ね順を逆に / Reverse stacking order ---
                for (var i = 0; i < generatedPieces.length; i++) {
                    generatedPieces[i].zOrder(ZOrderMethod.SENDTOBACK);
                }

                // --- 追加: アートボードに変換 / Convert to Artboards ---
                if (artboardCheckbox.value) {
                    // クリップグループ（generatedPieces）をマスク矩形の見た目順（Y降順, X昇順）で並べ替え
                    var sortItems = [];
                    for (var i = 0; i < generatedPieces.length; i++) {
                        var item = generatedPieces[i];
                        var maskRect = null;
                        if (item.typename === "GroupItem" && item.clipped) {
                            // クリッピングマスク（矩形）を取得
                            for (var j = 0; j < item.pageItems.length; j++) {
                                var pi = item.pageItems[j];
                                if (pi.clipping) {
                                    maskRect = pi;
                                    break;
                                }
                            }
                        } else if (item.typename === "PathItem") {
                            maskRect = item;
                        }
                        if (maskRect) {
                            var gb = maskRect.geometricBounds;
                            var left = gb[0];
                            var top = gb[1];
                            var right = gb[2];
                            var bottom = gb[3];
                            sortItems.push({
                                index: i,
                                item: item,
                                maskRect: maskRect,
                                left: left,
                                top: top,
                                right: right,
                                bottom: bottom
                            });
                        }
                    }
                    // Y座標降順（上から下）, 次にX座標昇順（左から右）
                    sortItems.sort(function(a, b) {
                        // Illustrator座標系は上が大きい
                        if (a.top != b.top) return b.top - a.top;
                        return a.left - b.left;
                    });
                    // アートボード追加 & リネーム（ゼロ埋め対応）/ Add artboards & rename (zero padding)
                    var doc = app.activeDocument;
                    var baseName = artboardNameInput.text;
                    var startNumberStr = artboardNumberInput.text;
                    var startNumber = parseInt(startNumberStr, 10);
                    if (isNaN(startNumber)) startNumber = 1;
                    // 総アートボード数・最大番号・ゼロ埋め桁数決定
                    var totalBoards = sortItems.length;
                    var maxNumber = startNumber + totalBoards - 1;
                    var digitCount = zeroPadCheckbox.value ? String(maxNumber).length : startNumberStr.length;
                    // マージン値取得
                    var marginVal = parseFloat(artboardMarginInput.text);
                    if (isNaN(marginVal)) marginVal = 0;
                    for (var k = 0; k < sortItems.length; k++) {
                        var rect = sortItems[k].maskRect;
                        var gb = rect.geometricBounds;
                        // マージン分拡張したアートボード座標
                        var left = gb[0] - marginVal;
                        var top = gb[1] + marginVal;
                        var right = gb[2] + marginVal;
                        var bottom = gb[3] - marginVal;
                        // IllustratorのaddArtboardは [left, top, right, bottom]
                        doc.artboards.add([left, top, right, bottom]);
                        // ゼロ埋め連番
                        var numStr = String(startNumber + k);
                        if (zeroPadCheckbox.value && numStr.length < digitCount) {
                            numStr = Array(digitCount - numStr.length + 1).join("0") + numStr;
                        }
                        var abName = baseName + numStr;
                        doc.artboards[doc.artboards.length - 1].name = abName;
                        // --- クリップグループの名前もアートボード名に設定 ---
                        var grp = sortItems[k].item;
                        if (grp && typeof grp.name !== "undefined") {
                            grp.name = abName;
                        }
                    }
                    // 元のアートボード削除 / Remove original artboards
                    for (var ai = doc.artboards.length - sortItems.length - 1; ai >= 0; ai--) {
                        doc.artboards.remove(ai);
                    }
                }

                // 元画像・一時矩形削除 / Remove original image and temp rect
                if (origImageObj && typeof origImageObj.remove === "function") {
                    origImageObj.remove();
                }
                if (isTempRect && selectedObjectForPuzzle && typeof selectedObjectForPuzzle.remove === "function") {
                    selectedObjectForPuzzle.remove();
                }
                return;
            }
        } else {
            Dialog.hide();
        }
    }, "スクリプト実行中");
}

/*
Offset Path Effect Utility
*/
/**
 * Offset PathエフェクトXMLを生成 / Generate Offset Path effect XML
 * @param {number} offsetVal
 * @return {string}
 */
function createOffsetEffectXML(offsetVal) {
    var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 2 "/></LiveEffect>';
    return xml.replace("value", offsetVal);
}

/**
 * 選択アイテムにオフセットパス効果を適用 / Apply Offset Path effect to selection
 * @param {number} offsetVal
 * @return {Object|null}
 */
function applyOffsetPathToSelection(offsetVal) {
    var result = null;
    safeExecute(function() {
        var doc = app.activeDocument;
        var original = doc.selection[0];
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        doc.selection = null;
        var copy = original.duplicate(original, ElementPlacement.PLACEAFTER);
        original.remove();
        copy.selected = true;
        copy.applyEffect(createOffsetEffectXML(offsetVal));
        app.redraw();
        app.executeMenuCommand('expandStyle');
        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        if (app.selection && app.selection.length > 0) {
            result = app.selection[0];
        } else {
            result = null;
        }
    }, "オフセットパスエフェクト適用");
    return result;
}

main();

/**
 * シンボル化できるオブジェクトのみ選別 / Filter symbolizable objects
 * @param {Array} items
 * @return {Array}
 */
function filterSymbolizableItems(items) {
    var validItems = [];
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item.locked && item.visible && !item.guides && !item.clipping) {
            validItems.push(item);
        }
    }
    return validItems;
}

/**
 * 選択オブジェクトをシンボル化し、必要なら矩形を生成して返す / Symbolize selected object(s), create rectangle if needed
 * @param {Array|Object} selection - app.selection
 * @return {Object} {symbolItem, maskRect, origImageObj, isTempRect}
 */
function convertToSymbol(selection) {
    var result = {
        symbolItem: null,
        maskRect: null,
        origImageObj: null,
        isTempRect: false
    };
    var doc = app.activeDocument;
    var selectedObj = null;

    // --- 複数選択時はグループ化してシンボル化（重ね順を保持） ---
    if (selection.length > 1) {
        try {
            var tempGroup = doc.groupItems.add();
            var selectedItems = [];
            for (var i = 0; i < selection.length; i++) {
                selectedItems.push(selection[i]);
            }
            // 後ろのものから先に移動することで重ね順を保持
            for (var j = selectedItems.length - 1; j >= 0; j--) {
                selectedItems[j].move(tempGroup, ElementPlacement.PLACEATBEGINNING);
            }
            var groupLeft = tempGroup.left;
            var groupTop = tempGroup.top;
            var groupParent = tempGroup.parent;
            var groupedSymbol = doc.symbols.add(tempGroup);
            var symbolItem = groupParent.symbolItems.add(groupedSymbol);
            symbolItem.left = groupLeft;
            symbolItem.top = groupTop;
            tempGroup.remove();
            selectedObj = symbolItem;
        } catch (e) {
            alert("複数オブジェクトのシンボル化に失敗しました: " + e);
            return result;
        }
    } else {
        selectedObj = selection[0];
    }

    // --- リンク画像の場合は複製してマスク、元のオブジェクトを削除 / Duplicate linked image and delete original ---
    if (selectedObj.typename === "PlacedItem" && !selectedObj.embedded) {
        var placedCopy = selectedObj.duplicate();
        selectedObj.remove(); // 元のリンクオブジェクト削除
        result.symbolItem = placedCopy;

        // 矩形作成
        var gb = placedCopy.geometricBounds;
        var rectWidth = gb[2] - gb[0];
        var rectHeight = gb[1] - gb[3];
        if (rectHeight < 0) rectHeight = -rectHeight;
        var rect = doc.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);

        result.maskRect = rect;
        result.origImageObj = placedCopy;
        result.isTempRect = true;

        return result;
    }

    // --- 埋め込み画像(RasterItem)をSymbolItemに変換 ---
    if (selectedObj.typename === "RasterItem") {
        try {
            var rasterLeft = selectedObj.left;
            var rasterTop = selectedObj.top;
            var rasterParent = selectedObj.parent;
            var tempSymbol = doc.symbols.add(selectedObj);
            var symbolItem = rasterParent.symbolItems.add(tempSymbol);
            symbolItem.left = rasterLeft;
            symbolItem.top = rasterTop;
            selectedObj.remove();
            selectedObj = symbolItem;
        } catch (e) {
            alert("埋め込み画像のシンボル化に失敗しました: " + e);
            return result;
        }
    }
    // --- ベクターオブジェクトをSymbolItemに変換 ---
    else if (
        selectedObj.typename === "PathItem" ||
        selectedObj.typename === "GroupItem" ||
        selectedObj.typename === "CompoundPathItem"
    ) {
        try {
            var vectorLeft = selectedObj.left;
            var vectorTop = selectedObj.top;
            var vectorParent = selectedObj.parent;
            var vectorSymbol = doc.symbols.add(selectedObj);
            var vectorSymbolItem = vectorParent.symbolItems.add(vectorSymbol);
            vectorSymbolItem.left = vectorLeft;
            vectorSymbolItem.top = vectorTop;
            selectedObj.remove();
            selectedObj = vectorSymbolItem;
        } catch (e) {
            alert("ベクターオブジェクトのシンボル化に失敗しました: " + e);
            return result;
        }
    }

    // PlacedItemまたはRasterItemまたはSymbolItemの場合は矩形化
    var selType = selectedObj.typename;
    if (selType == "PlacedItem" || selType == "RasterItem" || selType == "SymbolItem") {
        var gb = selectedObj.geometricBounds;
        var rectWidth = gb[2] - gb[0];
        var rectHeight = gb[1] - gb[3];
        if (rectHeight < 0) rectHeight = -rectHeight;
        var rect = doc.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);
        result.maskRect = rect;
        result.origImageObj = selectedObj;
        result.isTempRect = true;
    }
    result.symbolItem = selectedObj;
    return result;
}