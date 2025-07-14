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

*/

function getCurrentLang() {
    // 言語判定 / Language detection
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/* スクリプトバージョン / Script version */
var SCRIPT_VERSION = "v1.4";

/*
UIラベル定義 / UI Label Definitions
必要なキーのみを保持し、表示順序をUIに合わせる
*/
var LABELS = {
    dialogTitle: {
        ja: "アートボードサイズを調整 " + SCRIPT_VERSION,
        en: "Adjust Artboard Size " + SCRIPT_VERSION
    },
    targetSelection: {
        ja: "選択したオブジェクト",
        en: "Selected Objects"
    },
    targetArtboard: {
        ja: "現在のアートボード",
        en: "Current Artboard"
    },
    targetAllArtboards: {
        ja: "すべてのアートボード",
        en: "All Artboards"
    },
    marginLabel: {
        ja: "マージン",
        en: "Margin"
    },
    numberAlert: {
        ja: "数値を入力してください。",
        en: "Please enter a number."
    }
};

/*
マージンダイアログ表示 / Show margin input dialog with live preview
*/
/*
マージンダイアログ表示 / Show margin input dialog with live preview
*/
function showMarginDialog(defaultValue, unit, lang, artboardCount, hasSelection) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;

    /* 対象選択パネル / Target selection panel */
    var targetPanel = dlg.add("panel", undefined, lang === 'ja' ? "対象" : "Target");
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 10];

    var radioGroup = targetPanel.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var radioSelection = radioGroup.add("radiobutton", undefined, LABELS.targetSelection[lang]);
    radioSelection.enabled = hasSelection;
    var radioArtboard = radioGroup.add("radiobutton", undefined, LABELS.targetArtboard[lang]);
    var radioAllArtboards = radioGroup.add("radiobutton", undefined, LABELS.targetAllArtboards[lang]);

    // 対象選択の初期値設定 / Set initial radio selection
    if (!hasSelection && artboardCount === 1) {
        radioArtboard.value = true;
    } else if (!hasSelection && artboardCount > 1) {
        radioAllArtboards.value = true;
    } else {
        radioSelection.value = true;
    }

    /* マージン入力グループ / Margin input group */
    var inputSubGroup = dlg.add("group");
    inputSubGroup.orientation = "row";
    var label = inputSubGroup.add("statictext", undefined, LABELS.marginLabel[lang] + ":");
    var input = inputSubGroup.add("edittext", undefined, defaultValue);
    input.characters = 4;
    var unitLabel = inputSubGroup.add("statictext", undefined, unit);

    /* 現在のアートボードrectと全アートボードrectを保存（プレビュー用に復元） */
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }

    /*
    プレビュー更新関数 / Update artboard preview for dialog
    入力値・対象に応じてアートボードを一時的に調整
    */
    function updatePreview(value) {
        var previewValue = parseFloat(value);
        if (isNaN(previewValue)) return;
        var previewMarginInPoints = new UnitValue(previewValue, unit).as('pt');
        var targetMode = radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards");

        if (targetMode === "allArtboards") {
            for (var i = 0; i < app.activeDocument.artboards.length; i++) {
                var baseRect = originalRects[i].slice();
                var bounds = baseRect;
                bounds[0] -= previewMarginInPoints;
                bounds[1] += previewMarginInPoints;
                bounds[2] += previewMarginInPoints;
                bounds[3] -= previewMarginInPoints;
                app.activeDocument.artboards[i].artboardRect = bounds;
            }
            app.redraw();
            return;
        }
        if (targetMode === "artboard") {
            var baseRect = originalRects[abIndex].slice();
            var bounds = baseRect;
            bounds[0] -= previewMarginInPoints;
            bounds[1] += previewMarginInPoints;
            bounds[2] += previewMarginInPoints;
            bounds[3] -= previewMarginInPoints;
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
        previewBounds[0] -= previewMarginInPoints;
        previewBounds[1] += previewMarginInPoints;
        previewBounds[2] += previewMarginInPoints;
        previewBounds[3] -= previewMarginInPoints;
        app.activeDocument.artboards[abIndex].artboardRect = previewBounds;
        app.redraw();
    }

    // 入力欄で矢印キーによる増減を可能に / Enable arrow key increment/decrement in input
    changeValueByArrowKey(input, updatePreview);
    input.active = true;
    input.onChanging = function() {
        updatePreview(input.text);
    };
    radioSelection.onClick = function() { updatePreview(input.text); };
    radioArtboard.onClick = function() { updatePreview(input.text); };
    radioAllArtboards.onClick = function() { updatePreview(input.text); };

    /* ボタングループ / Button group */
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, lang === 'ja' ? "キャンセル" : "Cancel", { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, lang === 'ja' ? "OK" : "OK", { name: "ok" });
    btnGroup.margins = [0, 5, 0, 0];

    var result = null;
    okBtn.onClick = function() {
        if (!isNaN(parseFloat(input.text))) {
            result = {
                margin: input.text,
                target: radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards")
            };
            updatePreview(result.margin);
            dlg.close();
        } else {
            alert(LABELS.numberAlert[lang]);
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
    updatePreview(defaultValue);
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
    var lang = getCurrentLang();

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
        var userInput = showMarginDialog(defaultMarginValue, marginUnit, lang, artboards.length, doc.selection.length > 0);
        if (!userInput) return;

        /* allArtboards選択時は計算用マージンを0に / Set margin to 0 for allArtboards calculation */
        if (userInput.target === "allArtboards") {
            defaultMarginValue = '0';
        }

        marginValue = parseFloat(userInput.margin);
        var targetMode = userInput.target;
        marginInPoints = new UnitValue(marginValue, marginUnit).as('pt');

        if (targetMode === "artboard") {
            var artRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
            var bounds = artRect.slice();
            bounds[0] -= marginInPoints;
            bounds[1] += marginInPoints;
            bounds[2] += marginInPoints;
            bounds[3] -= marginInPoints;

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
                bounds[0] -= marginInPoints;
                bounds[1] += marginInPoints;
                bounds[2] += marginInPoints;
                bounds[3] -= marginInPoints;

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
            グループ内の clipping=true のみ抽出し直す
            Only add path with clipping=true in group
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
        selectedBounds[0] -= marginInPoints;
        selectedBounds[1] += marginInPoints;
        selectedBounds[2] += marginInPoints;
        selectedBounds[3] -= marginInPoints;

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
        alert(lang === 'ja' ? "エラーが発生しました: " + e.message : "An error occurred: " + e.message);
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
visibleBounds を常に使用
*/
function getBounds(item) {
    return item.visibleBounds;
}

/*
edittextに矢印キーで値を増減する機能を追加
Add arrow key increment/decrement to edittext
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
