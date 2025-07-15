#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

FitArtboardWithMargin.jsx

### 概要

- 選択オブジェクトまたはすべてのオブジェクトのバウンディングボックスにマージンを加え、アートボードを自動調整します。
- 定規単位に応じた初期マージン値と即時プレビュー付きダイアログを提供します。
- ピクセル整数値に丸めてアートボードを設定します。

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数丸め

### 処理の流れ

1. 対象（選択オブジェクトまたはアートボード）を選択
2. ダイアログでマージン値を設定（即時プレビュー対応）
3. 設定に基づきアートボードを自動調整

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に
- ↑↓キー、shift + ↑↓キーによる入力

### note

https://note.com/dtp_tranist/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.1 (20250708) : UI改善、ポイント初期値変更
- v1.2 (20250709) : UI改善とバグ修正
- v1.3 (20250710) : 「対象：現在のアートボード、すべてのアートボード」を追加
- v1.4 (20250713) : 矢印キーによる値変更機能を追加、UI改善
- v1.5 (20250715) : 上下・左右個別に設定できるように

---

### Script Name:

FitArtboardWithMargin.jsx

### Overview

- Automatically resize the artboard to fit the bounding box of selected or all objects with margin.
- Provides unit-based default margin values and an instant preview dialog.
- Sets the artboard size rounded to pixel integers.

### Main Features

- Default margin values based on ruler units
- Bounding box calculation
- Dialog with live preview
- Integer pixel rounding

### Workflow

1. Select target (selection or artboard)
2. Set margin value in dialog (with live preview)
3. Automatically adjust artboard size based on settings

### Changelog

- v1.0 (20250420): Initial version
- v1.1 (20250708): UI improvements, updated default point value
- v1.2 (20250709): UI improvements and bug fixes
- v1.3 (20250710): Added "Target: Current Artboard, All Artboards" options
- v1.4 (20250713): Added arrow key value change feature, UI improvements
- v1.5 (20250715): Enabled separate settings for vertical and horizontal margins

*/
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

var SCRIPT_VERSION = "v1.5";

var LABELS = {
    dialogTitle: { ja: "アートボードサイズを調整 " + SCRIPT_VERSION, en: "Adjust Artboard Size " + SCRIPT_VERSION },
    targetSelection: { ja: "選択したオブジェクト", en: "Selected Objects" },
    targetArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
    targetAllArtboards: { ja: "すべてのアートボード", en: "All Artboards" },
    marginLabel: { ja: "マージン", en: "Margin" },
    marginVertical: { ja: "上下", en: "Vertical" },
    marginHorizontal: { ja: "左右", en: "Horizontal" },
    linked: { ja: "連動", en: "Linked" },
    numberAlert: { ja: "数値を入力してください。", en: "Please enter a number." }
};

var offsetX = 300;
var dialogOpacity = 0.95;

/*
マージンダイアログ表示 / Show margin input dialog with live preview
*/
function showMarginDialog(defaultValue, unit, artboardCount, hasSelection) {
    var dlg = new Window("dialog", LABELS.dialogTitle);
    // --- dialog offset and opacity customization ---

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            dlg.layout.layout(true);
            var dialogWidth = dlg.bounds.width;
            var dialogHeight = dlg.bounds.height;

            var screenWidth = $.screens[0].right - $.screens[0].left;
            var screenHeight = $.screens[0].bottom - $.screens[0].top;

            var centerX = screenWidth / 2 - dialogWidth / 2;
            var centerY = screenHeight / 2 - dialogHeight / 2;

            dlg.location = [centerX + offsetX, centerY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, 0);
    // --- end dialog offset and opacity customization ---

    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;

    /* 対象選択パネル / Target selection panel */
    var targetPanel = dlg.add("panel", undefined, ($.locale.indexOf('ja') === 0 ? "対象" : "Target"));
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 10];

    var radioGroup = targetPanel.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var radioSelection = radioGroup.add("radiobutton", undefined, LABELS.targetSelection);
    radioSelection.enabled = hasSelection;
    var radioArtboard = radioGroup.add("radiobutton", undefined, LABELS.targetArtboard);
    var radioAllArtboards = radioGroup.add("radiobutton", undefined, LABELS.targetAllArtboards);

    // 対象選択の初期値設定 / Set initial radio selection
    if (!hasSelection && artboardCount === 1) {
        radioArtboard.value = true;
    } else if (!hasSelection && artboardCount > 1) {
        radioAllArtboards.value = true;
    } else {
        radioSelection.value = true;
    }

    /* マージン入力パネル / Margin input panel */
    var marginPanel = dlg.add("panel", undefined, (($.locale.indexOf('ja') === 0 ? "マージン" : "Margin") + " (" + unit + ")"));
    marginPanel.orientation = "row";
    marginPanel.alignChildren = ["fill", "top"];
    marginPanel.margins = [15, 20, 15, 10];

    var leftColumn = marginPanel.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["left", "center"];

    var rightColumn = marginPanel.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["left", "center"];
    rightColumn.alignment = ["fill", "center"];

    // 上下 group
    var verticalGroup = leftColumn.add("group");
    verticalGroup.orientation = "row";
    var labelV = verticalGroup.add("statictext", undefined, LABELS.marginVertical + ":");
    var inputV = verticalGroup.add("edittext", undefined, defaultValue);
    inputV.characters = 4;

    // 左右 group
    var horizontalGroup = leftColumn.add("group");
    horizontalGroup.orientation = "row";
    var labelH = horizontalGroup.add("statictext", undefined, LABELS.marginHorizontal + ":");
    var inputH = horizontalGroup.add("edittext", undefined, defaultValue);
    inputH.characters = 4;

    // 連動チェックボックス
    var linkCheckbox = rightColumn.add("checkbox", undefined, LABELS.linked);
    linkCheckbox.value = true;

    /* 現在のアートボードrectと全アートボードrectを保存（プレビュー用に復元） / Save current and all artboard rects for preview restore */
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }

    /*
    プレビュー更新関数 / Update artboard preview for dialog
    入力値・対象に応じてアートボードを一時的に調整 / Temporarily adjust artboard for preview
    */
    function updatePreview(valueV, valueH) {
        var previewValueV = parseFloat(valueV);
        var previewValueH = parseFloat(valueH);
        if (isNaN(previewValueV)) return;
        if (isNaN(previewValueH)) previewValueH = previewValueV;
        var previewMarginV = new UnitValue(previewValueV, unit).as('pt');
        var previewMarginH = new UnitValue(previewValueH, unit).as('pt');
        var targetMode = radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards");

        if (targetMode === "allArtboards") {
            for (var i = 0; i < app.activeDocument.artboards.length; i++) {
                var baseRect = originalRects[i].slice();
                var bounds = baseRect;
                bounds[0] -= previewMarginH;
                bounds[1] += previewMarginV;
                bounds[2] += previewMarginH;
                bounds[3] -= previewMarginV;
                app.activeDocument.artboards[i].artboardRect = bounds;
            }
            app.redraw();
            return;
        }
        if (targetMode === "artboard") {
            var baseRect = originalRects[abIndex].slice();
            var bounds = baseRect;
            bounds[0] -= previewMarginH;
            bounds[1] += previewMarginV;
            bounds[2] += previewMarginH;
            bounds[3] -= previewMarginV;
            app.activeDocument.artboards[abIndex].artboardRect = bounds;
            app.redraw();
            return;
        }
        // selectionモード
        var previewItems = app.activeDocument.selection.length === 0 ? app.activeDocument.pageItems : app.activeDocument.selection;
        var tempItems = [];
        for (var i = 0; i < previewItems.length; i++) {
            var item = previewItems[i];
            if (item.typename === "GroupItem" && item.clipped) {
                for (var j = 0; j < item.pageItems.length; j++) {
                    var child = item.pageItems[j];
                    if (child.clipping) {
                        tempItems.push(child);
                        break;
                    }
                }
            } else {
                tempItems.push(item);
            }
        }
        if (tempItems.length === 0) return;
        var previewBounds = getMaxBounds(tempItems);
        previewBounds[0] -= previewMarginH;
        previewBounds[1] += previewMarginV;
        previewBounds[2] += previewMarginH;
        previewBounds[3] -= previewMarginV;
        app.activeDocument.artboards[abIndex].artboardRect = previewBounds;
        app.redraw();
    }

    // 入力欄で矢印キーによる増減を可能に / Enable arrow key increment/decrement in input
    changeValueByArrowKey(inputV, function(val) {
        if (linkCheckbox.value) inputH.text = val;
        updatePreview(inputV.text, inputH.text);
    });
    changeValueByArrowKey(inputH, function(val) {
        if (linkCheckbox.value) inputV.text = val;
        updatePreview(inputV.text, inputH.text);
    });
    inputV.active = true;
    inputV.onChanging = function() {
        if (linkCheckbox.value) inputH.text = inputV.text;
        updatePreview(inputV.text, inputH.text);
    };
    inputH.onChanging = function() {
        if (linkCheckbox.value) inputV.text = inputH.text;
        updatePreview(inputV.text, inputH.text);
    };
    linkCheckbox.onClick = function() {
        if (linkCheckbox.value) inputH.text = inputV.text;
    };
    radioSelection.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };
    radioArtboard.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };
    radioAllArtboards.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };

    /* ボタングループ / Button group */
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, ($.locale.indexOf('ja') === 0 ? "キャンセル" : "Cancel"), {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, ($.locale.indexOf('ja') === 0 ? "OK" : "OK"), {
        name: "ok"
    });
    btnGroup.margins = [0, 5, 0, 0];

    var result = null;
    okBtn.onClick = function() {
        if (!isNaN(parseFloat(inputV.text)) && !isNaN(parseFloat(inputH.text))) {
            result = {
                marginV: inputV.text,
                marginH: inputH.text,
                target: radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards")
            };
            updatePreview(result.marginV, result.marginH);
            dlg.close();
        } else {
            alert(LABELS.numberAlert);
        }
    };
    cancelBtn.onClick = function() {
        /* プレビューで変更した全アートボードrectを元に戻す / Restore all artboard rects after preview */
        for (var i = 0; i < app.activeDocument.artboards.length; i++) {
            app.activeDocument.artboards[i].artboardRect = originalRects[i];
        }
        app.redraw();
        dlg.close();
    };
    updatePreview(inputV.text, inputH.text);
    dlg.show();
    return result;
}

/*
メイン処理 / Main process
*/
function main() {
    var selectedItems, artboards, rulerType, marginUnit, marginValue;
    var marginInPoints, defaultMarginValue, artboardIndex, selectedBounds;
    var supportedUnits = ['inch', 'mm', 'pt', 'pica', 'cm', 'H', 'px'];

    try {
        var doc = app.activeDocument;
        selectedItems = doc.selection;
        if (selectedItems.length === 0) {
            selectedItems = doc.pageItems;
            if (selectedItems.length === 0) return;
        }

        artboards = doc.artboards;
        rulerType = app.preferences.getIntegerPreference("rulerType");
        marginUnit = supportedUnits[rulerType];

        /* 単位ごとの初期マージン値設定 / Set default margin value based on unit */
        defaultMarginValue = '0';
        if (marginUnit === 'mm') {
            defaultMarginValue = '5';
        } else if (marginUnit === 'px') {
            defaultMarginValue = '20';
        } else if (marginUnit === 'pt') {
            defaultMarginValue = '10';
        }

        /* 選択なし・複数アートボード時は allArtboards をデフォルトに / Default to allArtboards if no selection and multiple artboards */
        var isAllArtboardsDefault = (doc.selection.length === 0 && artboards.length > 1);
        if (isAllArtboardsDefault) {
            defaultMarginValue = '0';
        }

        /* ユーザーにマージンを入力させる / Show margin input dialog */
        var userInput = showMarginDialog(defaultMarginValue, marginUnit, artboards.length, doc.selection.length > 0);
        if (!userInput) return;

        /* allArtboards選択時は計算用マージンを0に / Set margin to 0 for allArtboards calculation */
        if (userInput.target === "allArtboards") {
            defaultMarginValue = '0';
        }

        // 新しいUI: userInput.marginV, userInput.marginH
        var marginV = parseFloat(userInput.marginV);
        var marginH = parseFloat(userInput.marginH);
        var targetMode = userInput.target;
        var marginVInPoints = new UnitValue(marginV, marginUnit).as('pt');
        var marginHInPoints = new UnitValue(marginH, marginUnit).as('pt');

        if (targetMode === "artboard") {
            var artRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
            var bounds = artRect.slice();
            bounds[0] -= marginHInPoints;
            bounds[1] += marginVInPoints;
            bounds[2] += marginHInPoints;
            bounds[3] -= marginVInPoints;

            var x0 = Math.round(bounds[0]);
            var y1 = Math.round(bounds[1]);
            var x2 = Math.round(bounds[2]);
            var y3 = Math.round(bounds[3]);

            var width = Math.round(x2 - x0);
            var height = Math.round(y1 - y3);

            x2 = x0 + width;
            y3 = y1 - height;

            bounds[0] = x0;
            bounds[1] = y1;
            bounds[2] = x2;
            bounds[3] = y3;

            artboards[artboards.getActiveArtboardIndex()].artboardRect = bounds;
            return;
        }

        if (targetMode === "allArtboards") {
            for (var i = 0; i < artboards.length; i++) {
                var artRect = artboards[i].artboardRect;
                var bounds = artRect.slice();
                bounds[0] -= marginHInPoints;
                bounds[1] += marginVInPoints;
                bounds[2] += marginHInPoints;
                bounds[3] -= marginVInPoints;

                var x0 = Math.round(bounds[0]);
                var y1 = Math.round(bounds[1]);
                var x2 = Math.round(bounds[2]);
                var y3 = Math.round(bounds[3]);

                var width = Math.round(x2 - x0);
                var height = Math.round(y1 - y3);

                x2 = x0 + width;
                y3 = y1 - height;

                bounds[0] = x0;
                bounds[1] = y1;
                bounds[2] = x2;
                bounds[3] = y3;

                artboards[i].artboardRect = bounds;
            }
            return;
        }

        if (targetMode === "selection") {
            /*
            グループ内の clipping=true のみ抽出し直す / Only add path with clipping=true in group
            */
            var tempItems = [];
            for (var i = 0; i < selectedItems.length; i++) {
                var item = selectedItems[i];
                if (item.typename === "GroupItem" && item.clipped) {
                    for (var j = 0; j < item.pageItems.length; j++) {
                        var child = item.pageItems[j];
                        if (child.clipping) {
                            tempItems.push(child);
                            break; // 最初に見つけた clipping=true を追加後に終了
                        }
                    }
                } else {
                    tempItems.push(item);
                }
            }
            selectedItems = tempItems;
            if (selectedItems.length === 0) return;
        }

        /* 選択範囲のバウンディングボックス計算 / Calculate bounding box of selected items */
        selectedBounds = getMaxBounds(selectedItems);
        selectedBounds[0] -= marginHInPoints;
        selectedBounds[1] += marginVInPoints;
        selectedBounds[2] += marginHInPoints;
        selectedBounds[3] -= marginVInPoints;

        /* 座標と幅・高さを整数に丸める / Round coordinates and size to integers */
        var x0 = Math.round(selectedBounds[0]);
        var y1 = Math.round(selectedBounds[1]);
        var x2 = Math.round(selectedBounds[2]);
        var y3 = Math.round(selectedBounds[3]);

        var width = Math.round(x2 - x0);
        var height = Math.round(y1 - y3);

        x2 = x0 + width;
        y3 = y1 - height;

        selectedBounds[0] = x0;
        selectedBounds[1] = y1;
        selectedBounds[2] = x2;
        selectedBounds[3] = y3;

        /* アートボードの更新 / Update artboard */
        artboardIndex = artboards.getActiveArtboardIndex();
        artboards[artboardIndex].artboardRect = selectedBounds;

    } catch (e) {
        alert(($.locale.indexOf('ja') === 0 ? "エラーが発生しました: " : "An error occurred: ") + e.message);
    }
}

/*
選択オブジェクト群から最大のバウンディングボックスを取得 / Get maximum bounding box from multiple items
*/
function getMaxBounds(items) {
    var bounds = getBounds(items[0]);
    for (var i = 1; i < items.length; i++) {
        var itemBounds = getBounds(items[i]);
        bounds[0] = Math.min(bounds[0], itemBounds[0]);
        bounds[1] = Math.max(bounds[1], itemBounds[1]);
        bounds[2] = Math.max(bounds[2], itemBounds[2]);
        bounds[3] = Math.min(bounds[3], itemBounds[3]);
    }
    return bounds;
}

/*
オブジェクトのバウンディングボックスを取得 / Get bounding box of a single object
visibleBounds を常に使用 / Always use visibleBounds
*/
function getBounds(item) {
    return item.visibleBounds;
}

/*
edittextに矢印キーで値を増減する機能を追加 / Add arrow key increment/decrement to edittext
*/
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (keyboard.shiftKey) {
                // 10の倍数にスナップ
                var base = Math.round(value / 10) * 10;
                if (event.keyName == "Up") {
                    value = base + 10;
                } else {
                    value = base - 10;
                    if (value < 0) value = 0; // 負数を防ぐ（必要なら）
                }
            } else {
                var delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") {
                onUpdate(editText.text);
            }
        }
    });
}

main();
app.selectTool("Adobe Select Tool");