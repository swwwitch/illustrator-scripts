#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ResizeClipMask.jsx

### 概要

- クリップグループ内のマスクパスを自動検出し、マージンを調整した新しいマスクに置き換えます。
- 複数クリップグループに一括適用可能で、長方形マスクのみ対応。

### 主な機能

- マスクパス検出と選択
- ユーザー指定マージンのダイアログ入力
- 正負切替ボタンによる値反転
- 長方形判定とスキップ処理

### 処理の流れ

1. クリップグループを選択
2. ダイアログでマージンを指定
3. 長方形マスクを検出し、新しいマスクに置換
4. 元のマスクパスを削除

### 更新履歴

- v1.0.0 (20250710) : 初期バージョン

---

### Script Name:

ResizeClipMask.jsx

### Overview

- Automatically detects mask paths inside clip groups and replaces them with new masks adjusted by user-defined margins.
- Supports batch processing of multiple clip groups (rectangular masks only).

### Main Features

- Detect and select mask paths
- Margin input via dialog
- Sign toggle button for margin inversion
- Rectangle detection and skip logic

### Workflow

1. Select clip groups
2. Specify margin in dialog
3. Detect rectangular masks and replace with new mask
4. Delete original mask path

### Change Log

- v1.0.0 (20250710): Initial release
*/

// スクリプトバージョン / Script version
var SCRIPT_VERSION = "v1.0";

// UIラベル定義（未使用項目は省略、出現順に整理）/ UI label definitions (unused items omitted, ordered as in UI)
var LABELS = {
    notRect: { ja: "このマスクは長方形ではありません。処理をスキップします。", en: "This mask is not a rectangle. Skipping." },
    noMask: { ja: "マスクパスが見つかりませんでした。", en: "No mask path found." },
    noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    dialogTitle: {
        ja: "マスクパスのサイズ変更 " + SCRIPT_VERSION,
        en: "Resize Mask Path " + SCRIPT_VERSION
    },
    margin: { ja: "マージン", en: "Margin" }
};

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

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

function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// 長方形判定 / Check if path is rectangle
function isRectangle(pathItem) {
    if (pathItem.pathPoints.length != 4) {
        return false;
    }
    for (var i = 0; i < 4; i++) {
        var pt = pathItem.pathPoints[i];
        if (pt.pointType != PointType.CORNER) {
            return false;
        }
        if (!pointsEqual(pt.anchor, pt.leftDirection) || !pointsEqual(pt.anchor, pt.rightDirection)) {
            return false;
        }
    }
    return true;
}

// 2点が等しいか判定 / Check if two points are equal
function pointsEqual(p1, p2) {
    return p1[0] == p2[0] && p1[1] == p2[1];
}

// 増減・反転ボタンのイベント処理 / Button event handlers for plus, minus, swap
function handlePlus(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    input.text = String(val + 1);
}
function handleMinus(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    input.text = String(val - 1);
}
function handleSwap(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    input.text = String(val * -1);
}

// マージンダイアログを表示 / Show margin dialog
function showMarginDialog(defaultValue, unitLabel) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";
    dlg.margins = 15;

    var inputGroup = dlg.add("group");
    inputGroup.add("statictext", undefined, LABELS.margin[lang] + " (" + unitLabel + "):");

    var inputSubGroup = inputGroup.add("group");
    inputSubGroup.orientation = "row";

    var input = inputSubGroup.add("edittext", undefined, defaultValue);
    input.characters = 4;

    var buttonGroup = inputSubGroup.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.spacing = 1;

    var plusMinusGroup = buttonGroup.add("group");
    plusMinusGroup.orientation = "column";
    plusMinusGroup.spacing = 1;

    var plusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "+");
    var minusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "-");

    var zeroGroup = buttonGroup.add("group");
    zeroGroup.orientation = "column";
    zeroGroup.alignChildren = ["center", "center"];

    var swapBtn = zeroGroup.add("button", [0, 0, 20, 31], "±");

    plusBtn.onClick = function() { handlePlus(input); };
    minusBtn.onClick = function() { handleMinus(input); };
    swapBtn.onClick = function() { handleSwap(input); };

    var btns = dlg.add("group");
    btns.alignment = "right";
    var cancel = btns.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var ok = btns.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    input.active = true;
    var result = dlg.show();
    if (result != 1) return null;

    var margin = parseFloat(input.text);
    if (isNaN(margin)) {
        alert(LABELS.noSelection[lang]);
        return null;
    }
    return margin;
}

// 選択からマスクパスを抽出 / Collect mask paths from selection
function collectMaskPaths(selection) {
    var masks = [];
    for (var i = 0; i < selection.length; i++) {
        var group = selection[i];
        if (group.typename === "GroupItem" && group.clipped) {
            for (var j = 0; j < group.pageItems.length; j++) {
                var item = group.pageItems[j];
                if (item.typename === "PathItem" && item.clipping) {
                    masks.push(item);
                    break;
                }
            }
        }
    }
    return masks;
}

// マスク矩形の座標情報取得 / Get rectangle info from mask
function getRectInfo(mask) {
    return {
        left: mask.left,
        top: mask.top,
        width: mask.width,
        height: mask.height
    };
}

// マスクを置換 / Replace mask with new rectangle
function replaceMask(mask, margin) {
    if (!isRectangle(mask)) {
        alert(LABELS.notRect[lang]);
        return;
    }
    var parentGroup = mask.parent;
    var rect = getRectInfo(mask);
    var newRect = parentGroup.pathItems.rectangle(
        rect.top + margin,
        rect.left - margin,
        rect.width + (margin * 2),
        rect.height + (margin * 2)
    );
    newRect.filled = false;
    newRect.stroked = false;
    newRect.clipping = true;
    newRect.move(parentGroup, ElementPlacement.PLACEATBEGINNING);
    mask.remove();
    parentGroup.selected = true;
}

function main() {
    // ドキュメントが開かれていない場合 / No document open
    if (app.documents.length === 0) {
        alert(LABELS.noSelection[lang]);
        return;
    }
    var sel = app.activeDocument.selection;
    // 選択がない場合 / No selection
    if (sel.length === 0) {
        alert(LABELS.noSelection[lang]);
        return;
    }
    var marginUnit = getCurrentUnitLabel();
    var defaultMarginValue = '0';
    if (marginUnit === 'mm') {
        defaultMarginValue = '-5';
    } else if (marginUnit === 'px') {
        defaultMarginValue = '-20';
    } else if (marginUnit === 'pt') {
        defaultMarginValue = '-10';
    }
    // マージンダイアログを表示 / Show margin dialog
    var margin = showMarginDialog(defaultMarginValue, marginUnit);
    if (margin === null) return;
    // マスクパスを収集 / Collect mask paths
    var newSelection = collectMaskPaths(sel);
    if (newSelection.length > 0) {
        for (var i = 0; i < newSelection.length; i++) {
            replaceMask(newSelection[i], margin);
        }
    } else {
        alert(LABELS.noMask[lang]);
    }
}

main();