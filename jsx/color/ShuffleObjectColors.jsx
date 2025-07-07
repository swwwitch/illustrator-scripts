#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ShuffleObjectColors.jsx

### 概要

- 選択オブジェクト（パス、グループ、複合シェイプ）の塗り・線カラーを再配色します。
- 黒・白の除外、カラー比率（バランス）保持、ランダム適用が可能です。
- 日本語／英語UIに対応。

### 主な機能

- 塗り・線の適用切替
- 黒・白の除外
- カラー比率保持オプション
- ランダム／順番適用切替

### 処理の流れ

1. 対象オブジェクトを収集
2. 使用カラーを抽出
3. プールを生成・並べ替え
4. カラーを再適用

### 更新履歴

- v1.0.0 (20240624) : 初期バージョン
- v1.0.1 (20240624) : バグ修正
- v1.0.2 (20240624) : ランダム／順番適用切替追加
- v1.0.3 (20240625) : ローカライズ調整
- v1.1 (20250708) : テキスト機能削除、構造整理

---

### Script Name:

ShuffleObjectColors.jsx

### Overview

- Reapplies fill and stroke colors to selected objects (paths, groups, compound shapes).
- Supports excluding black/white, preserving color balance, and random application.
- Supports Japanese and English UI.

### Main Features

- Toggle fill and stroke application
- Exclude black and white
- Preserve color balance option
- Random or sequential color application

### Workflow

1. Collect target objects
2. Extract used colors
3. Create and shuffle color pool
4. Reapply colors

### Changelog

- v1.0.0 (20240624): Initial release
- v1.0.1 (20240624): Bug fixes
- v1.0.2 (20240624): Added random/sequential application
- v1.0.3 (20240625): Localization updates
- v1.1 (20250708): Removed text feature, refactored structure
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/* ラベル定義（UI出現順） */
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "カラーをランダムに再適用 v1.1", en: "Reapply Colors Randomly v1.1" },
    panelTarget: { ja: "対象", en: "Target" },
    fill: { ja: "塗り", en: "Fill" },
    stroke: { ja: "線", en: "Stroke" },
    panelExclude: { ja: "保護", en: "Exclude" },
    black: { ja: "黒", en: "Black" },
    white: { ja: "白", en: "White" },
    random: { ja: "ランダム", en: "Random" },
    balance: { ja: "バランスを保持", en: "Preserve Balance" },
    apply: { ja: "適用", en: "Apply" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select some objects." },
    alertEmpty: { ja: "使用可能なカラーが見つかりません（白と黒以外）。", en: "No usable colors found (excluding white and black)." },
    alertChoice: { ja: "塗りまたは線のいずれか一方を選択してください。", en: "Please select either fill or stroke." }
};

// UIパーツ
var fillCheckbox, strokeCheckbox;
var blackCheckbox, whiteCheckbox;
var balanceCheckbox;
var randomCheckbox;

main();

// 塗り・線のカラー適用
function applyColorIfEligible(target, isFill, changeFlag, colorArray, colorIndex, excludeBlack, excludeWhite) {
    if (!changeFlag) return { color: null, index: colorIndex };

    var currentColor = isFill ? target.fillColor : target.strokeColor;
    var unified = getOriginalColor(currentColor);
    if (!unified || (excludeBlack && isBlack(unified)) || (excludeWhite && isWhite(unified))) return { color: null, index: colorIndex };

    var nextColorIndex = colorIndex % colorArray.length;
    var nextColor = getOriginalColor(colorArray[nextColorIndex]);

    if (isFill) {
        target.filled = true;
        target.fillColor = nextColor;
    } else {
        target.stroked = true;
        target.strokeColor = nextColor;
    }
    colorIndex++;
    return { color: nextColor, index: colorIndex };
}

/* メイン処理 */
function main() {
    if (app.documents.length === 0 || !app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert(LABELS.alertSelect[lang]);
        return;
    }

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "row";
    dialog.alignChildren = "top";

    var leftGroup = dialog.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = "left";

    var optionPanel = leftGroup.add("panel", undefined, LABELS.panelTarget[lang]);
    optionPanel.preferredSize.width = 120;
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.margins= [10, 25, 10, 10];

    var checkGroup = optionPanel.add("group");
    checkGroup.orientation = "row";
    checkGroup.alignChildren = "left";

    fillCheckbox = checkGroup.add("checkbox", undefined, LABELS.fill[lang]);
    fillCheckbox.value = true;
    strokeCheckbox = checkGroup.add("checkbox", undefined, LABELS.stroke[lang]);
    strokeCheckbox.value = true;

    var excludePanel = leftGroup.add("panel", undefined, LABELS.panelExclude[lang]);
    excludePanel.preferredSize.width = 120;
    excludePanel.orientation = "column";
    excludePanel.alignChildren = "left";
    excludePanel.margins = [10, 25, 10, 10];

    var excludeGroup = excludePanel.add("group");
    excludeGroup.orientation = "row";
    excludeGroup.alignChildren = "left";

    blackCheckbox = excludeGroup.add("checkbox", undefined, LABELS.black[lang]);
    blackCheckbox.value = true;

    whiteCheckbox = excludeGroup.add("checkbox", undefined, LABELS.white[lang]);
    whiteCheckbox.value = true;

    randomCheckbox = leftGroup.add("checkbox", undefined, LABELS.random[lang]);
    randomCheckbox.value = true;

    balanceCheckbox = leftGroup.add("checkbox", undefined, LABELS.balance[lang]);
    balanceCheckbox.value = true;

    var rightGroup = dialog.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = "right";

    var okBtn = rightGroup.add("button", undefined, "OK");
    okBtn.preferredSize.width = 80;

    var applyBtn = rightGroup.add("button", undefined, LABELS.apply[lang]);
    applyBtn.preferredSize.width = 80;
    applyBtn.onClick = function() {
        applyBtn.enabled = false;
        applyColors();
        app.redraw();
        applyBtn.enabled = true;
    };

    var spacer = rightGroup.add("statictext", undefined, "");
    spacer.characters = 1;
    spacer.preferredSize.height = 50;

    var cancelBtn = rightGroup.add("button", undefined, LABELS.cancel[lang]);
    cancelBtn.preferredSize.width = 80;

    cancelBtn.onClick = function() {
        app.undo();
        dialog.close(0);
    };
    okBtn.onClick = function() {
        applyColors();
        dialog.close(1);
    };

    dialog.show();
}

function applyColors() {
    var changeFill = fillCheckbox.value;
    var changeStroke = strokeCheckbox.value;
    var excludeBlack = blackCheckbox.value;
    var excludeWhite = whiteCheckbox.value;
    var balancePreserve = balanceCheckbox.value;
    var randomize = (typeof randomCheckbox !== "undefined") ? randomCheckbox.value : true;

    if (!changeFill && !changeStroke) {
        alert(LABELS.alertChoice[lang]);
        return;
    }

    var selection = app.activeDocument.selection;
    var targetItems = [];
    collectAllTargetItems(selection, targetItems);

    var colorCountMap = {};
    var uniqueColorKeys = {};

    // カラー収集
    for (var i = 0; i < targetItems.length; i++) {
        var item = targetItems[i];

        if (changeFill && (item.filled)) {
            var fillColor = getOriginalColor(item.fillColor);
            if (
                fillColor &&
                !(excludeBlack && isBlack(fillColor)) &&
                !(excludeWhite && isWhite(fillColor))
            ) {
                collectColor(fillColor, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }

        if (changeStroke && (item.stroked) && item.strokeWidth >= 1) {
            var strokeColor = getOriginalColor(item.strokeColor);
            if (
                strokeColor &&
                !(excludeBlack && isBlack(strokeColor)) &&
                !(excludeWhite && isWhite(strokeColor))
            ) {
                collectColor(strokeColor, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }
    }

    var colorPool = createColorPool(colorCountMap, uniqueColorKeys, balancePreserve);

    if (colorPool.length === 0) {
        alert(LABELS.alertEmpty[lang]);
        return;
    }

    var shuffledColors = randomize ? shuffleArray(colorPool) : colorPool;
    var colorIndex = 0;

    // カラー再適用
    for (var j = 0; j < targetItems.length; j++) {
        var target = targetItems[j];
        var result = applyColorIfEligible(target, true, changeFill, shuffledColors, colorIndex, excludeBlack, excludeWhite);
        colorIndex = result.index;
        result = applyColorIfEligible(target, false, changeStroke, shuffledColors, colorIndex, excludeBlack, excludeWhite);
        colorIndex = result.index;
    }
}

function createColorPool(colorCountMap, uniqueColorKeys, balancePreserve) {
    var pool = [];
    if (balancePreserve) {
        for (var key in colorCountMap) {
            var entry = colorCountMap[key];
            for (var r = 0; r < entry.count; r++) {
                pool.push(getOriginalColor(entry.color));
            }
        }
    } else {
        for (var key in uniqueColorKeys) {
            pool.push(getOriginalColor(uniqueColorKeys[key]));
        }
    }
    return pool;
}

// カラー収集
function collectColor(color, colorCountMap, uniqueColorKeys, balancePreserve) {
    var key;
    if (color.typename === "RGBColor") {
        key = color.red + "," + color.green + "," + color.blue;
    } else if (color.typename === "CMYKColor") {
        key = color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black;
    } else {
        return;
    }
    if (!uniqueColorKeys[key]) {
        uniqueColorKeys[key] = getOriginalColor(color);
    }
    if (balancePreserve) {
        if (!colorCountMap[key]) {
            colorCountMap[key] = { color: getOriginalColor(color), count: 0 };
        }
        colorCountMap[key].count++;
    }
}

// 対象アイテム収集（再帰）
function collectAllTargetItems(items, results) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "GroupItem") {
            collectAllTargetItems(item.pageItems, results);
        } else if (item.typename === "CompoundPathItem") {
            collectAllTargetItems(item.pathItems, results);
        } else if (item.typename === "PathItem") {
            results.push(item);
        }
    }
}

// 配列シャッフル（Fisher–Yates）
function shuffleArray(arr) {
    var copy = arr.slice();
    for (var i = copy.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = copy[i];
        copy[i] = copy[j];
        copy[j] = temp;
    }
    return copy;
}

// カラー複製（RGB/CMYK対応）
function getOriginalColor(color) {
    if (!color || !color.typename) return null;
    if (color.typename === "RGBColor") return copyRGBColor(color);
    if (color.typename === "CMYKColor") return copyCMYKColor(color);
    return null;
}

function copyCMYKColor(original) {
    var newColor = new CMYKColor();
    newColor.cyan = original.cyan;
    newColor.magenta = original.magenta;
    newColor.yellow = original.yellow;
    newColor.black = original.black;
    return newColor;
}

function copyRGBColor(original) {
    var newColor = new RGBColor();
    newColor.red = original.red;
    newColor.green = original.green;
    newColor.blue = original.blue;
    return newColor;
}

// 黒判定
function isBlack(color) {
    if (color.typename === "RGBColor") {
        return color.red <= 51 && color.green <= 51 && color.blue <= 51;
    } else if (color.typename === "CMYKColor") {
        return color.black >= 0.8 || (color.cyan <= 0.2 && color.magenta <= 0.2 && color.yellow <= 0.2 && color.black >= 0.7);
    }
    return false;
}

// 白判定
function isWhite(color) {
    if (color.typename === "RGBColor") {
        return color.red === 255 && color.green === 255 && color.blue === 255;
    } else if (color.typename === "CMYKColor") {
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }
    return false;
}