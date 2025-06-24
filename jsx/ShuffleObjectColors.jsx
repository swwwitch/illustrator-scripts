#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
カラーランダム適用スクリプト for Adobe Illustrator

【概要】
選択オブジェクト（パス／テキスト／複合シェイプ／グループ）に対して、
塗り／線カラーをランダムに再適用します。
除外条件（黒・白）や、カラーの使用比率を保つオプションも搭載。

【処理の流れ】
1. 対象オブジェクトを収集（グループ・複合パス含む）
2. 指定条件に基づいてカラーを収集
3. カラーをシャッフル
4. 対象オブジェクトに塗り／線を再適用

【対象】
- 選択中の PathItem / TextFrame / GroupItem / CompoundPathItem

【限定条件】
- RGB / CMYK カラーモード対応（内部的に RGB に変換して処理）

更新日：2025-06-24
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// 日英ラベル定義 Define label（UI出現順に整理）
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "カラーをランダムに再適用", en: "Reapply Colors Randomly" },
    panelTarget: { ja: "対象", en: "Target" },
    fill: { ja: "塗り", en: "Fill" },
    stroke: { ja: "線", en: "Stroke" },
    panelExclude: { ja: "保護", en: "Exclude" },
    black: { ja: "黒", en: "Black" },
    white: { ja: "白", en: "White" },
    balance: { ja: "バランスを保持", en: "Preserve Balance" },
    ok: { ja: "OK", en: "OK" },
    apply: { ja: "適用", en: "Apply" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select some objects." },
    alertEmpty: { ja: "使用可能なカラーが見つかりません（白と黒以外）。", en: "No usable colors found (excluding white and black)." },
    alertChoice: { ja: "塗りまたは線のいずれか一方を選択してください。", en: "Please select either fill or stroke." }
};

// グローバル変数としてUIパーツを宣言
var fillCheckbox, strokeCheckbox;
var blackCheckbox, whiteCheckbox;
var balanceCheckbox; // ←追加

main();

// 塗り・線のカラー適用を共通化
function applyColorIfEligible(target, isFill, changeFlag, colorArray, colorIndex, otherColor, includeBlack, excludeWhite) {
    if (!changeFlag) return { color: null, index: colorIndex };

    var currentColor = isFill ? target.fillColor : target.strokeColor;
    var rgb = getUnifiedRGBColor(currentColor);
    if (!rgb || isProtectedColor(rgb, includeBlack, excludeWhite)) return { color: null, index: colorIndex };

    var nextColorIndex = colorIndex % colorArray.length;
    var nextColor = copyRGBColor(colorArray[nextColorIndex]);

    var attempts = 0;
    while (otherColor && colorsEqual(nextColor, otherColor) && attempts < colorArray.length) {
        colorIndex++;
        nextColorIndex = colorIndex % colorArray.length;
        nextColor = copyRGBColor(colorArray[nextColorIndex]);
        attempts++;
    }

    if (isFill) {
        target.fillColor = nextColor;
    } else {
        target.strokeColor = nextColor;
    }
    colorIndex++;
    return { color: nextColor, index: colorIndex };
}

function main() {
    // CMYK対応済みのためRGB限定チェックは削除
    /*
    if (app.activeDocument.documentColorSpace !== DocumentColorSpace.RGB) {
        alert(LABELS.alertRGB[lang]);
        return;
    }
    */

    if (app.documents.length === 0 || !app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert(LABELS.alertSelect[lang]);
        return;
    }

    // ダイアログ表示 / Show dialog
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "row";
    dialog.alignChildren = "top";

    // 左側のオプショングループ
    var leftGroup = dialog.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = "left";

    // パネル（カラー適用対象）/ Panel (Target for color application)
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

    // 除外条件パネル / Exclude panel
    var excludePanel = leftGroup.add("panel", undefined, LABELS.panelExclude[lang]);
    excludePanel.preferredSize.width = 120;
    excludePanel.orientation = "column";
    excludePanel.alignChildren = "left";
    excludePanel.margins = [10, 25, 10, 10];

    var excludeGroup = excludePanel.add("group");
    excludeGroup.orientation = "row";
    excludeGroup.alignChildren = "left";

    blackCheckbox = excludeGroup.add("checkbox", undefined, LABELS.black[lang]);
    blackCheckbox.value = false;

    whiteCheckbox = excludeGroup.add("checkbox", undefined, LABELS.white[lang]);
    whiteCheckbox.value = true;

    // 「バランスを保持」チェックボックス（パネル外・左下に追加） / "Preserve Balance" checkbox (outside panel, bottom left)
    balanceCheckbox = leftGroup.add("checkbox", undefined, LABELS.balance[lang]);
    balanceCheckbox.value = false;

    // 右側のボタングループ（縦並び） / Right side button group (vertical)
    var rightGroup = dialog.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = "right";

    var okBtn = rightGroup.add("button", undefined, LABELS.ok[lang]);
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
        dialog.close(0);
    };
    okBtn.onClick = function() {
        applyColors();
        dialog.close(1);
    };

    if (dialog.show() !== 1) return;

    // ここでの処理はapplyColors()に移動したため不要
}

function applyColors() {
    var changeFill = fillCheckbox.value;
    var changeStroke = strokeCheckbox.value;
    var includeBlack = blackCheckbox.value;
    var excludeWhite = whiteCheckbox.value;
    var balancePreserve = balanceCheckbox.value;

    if (!changeFill && !changeStroke) {
        alert(LABELS.alertChoice[lang]);
        return;
    }

    var selection = app.activeDocument.selection;
    var targetItems = [];
    collectAllTargetItems(selection, targetItems);

    var colorCountMap = {};
    var uniqueColorKeys = {};

    // カラー収集処理（共通）
    for (var i = 0; i < targetItems.length; i++) {
        var item = targetItems[i];

        if (changeFill && item.filled) {
            var rgbFill = getUnifiedRGBColor(item.fillColor);
            if (rgbFill && !isProtectedColor(rgbFill, includeBlack, excludeWhite)) {
                collectColor(rgbFill, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }

        if (changeStroke && item.stroked && item.strokeWidth >= 1) {
            var rgbStroke = getUnifiedRGBColor(item.strokeColor);
            if (rgbStroke && !isProtectedColor(rgbStroke, includeBlack, excludeWhite)) {
                collectColor(rgbStroke, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }
    }

    var colorPool = createColorPool(colorCountMap, uniqueColorKeys, balancePreserve);

    if (colorPool.length === 0) {
        alert(LABELS.alertEmpty[lang]);
        return;
    }

    var shuffledColors = shuffleArray(colorPool);
    var colorIndex = 0;

    // カラー再適用
    for (var j = 0; j < targetItems.length; j++) {
        var target = targetItems[j];
        var result = applyColorIfEligible(target, true, changeFill, shuffledColors, colorIndex, null, includeBlack, excludeWhite);
        colorIndex = result.index;
        result = applyColorIfEligible(target, false, changeStroke, shuffledColors, colorIndex, result.color, includeBlack, excludeWhite);
        colorIndex = result.index;
    }
}

function createColorPool(colorCountMap, uniqueColorKeys, balancePreserve) {
    var pool = [];
    if (balancePreserve) {
        for (var key in colorCountMap) {
            var entry = colorCountMap[key];
            for (var r = 0; r < entry.count; r++) {
                pool.push(copyRGBColor(entry.color));
            }
        }
    } else {
        for (var key in uniqueColorKeys) {
            pool.push(copyRGBColor(uniqueColorKeys[key]));
        }
    }
    return pool;
}

// カラー収集処理（共通）
function collectColor(rgb, colorCountMap, uniqueColorKeys, balancePreserve) {
    var key = rgb.red + "," + rgb.green + "," + rgb.blue;
    if (balancePreserve) {
        if (!colorCountMap[key]) {
            colorCountMap[key] = { color: copyRGBColor(rgb), count: 0 };
        }
        colorCountMap[key].count++;
    } else if (!uniqueColorKeys[key]) {
        uniqueColorKeys[key] = copyRGBColor(rgb);
    }
}

// 再帰的に PathItem, TextFrame を収集
function collectAllTargetItems(items, results) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "GroupItem") {
            collectAllTargetItems(item.pageItems, results);
        } else if (item.typename === "CompoundPathItem") {
            collectAllTargetItems(item.pathItems, results);
        } else if (item.typename === "PathItem" || item.typename === "TextFrame") {
            results.push(item);
        }
    }
}

// 配列をシャッフル（Fisher–Yatesアルゴリズム）
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

// 任意のカラーをRGBColorに変換（CMYK対応）
function getUnifiedRGBColor(color) {
    if (!color || !color.typename) return null;
    if (color.typename === "RGBColor") {
        return copyRGBColor(color);
    } else if (color.typename === "CMYKColor") {
        var c = color.cyan / 100;
        var m = color.magenta / 100;
        var y = color.yellow / 100;
        var k = color.black / 100;
        var rgb = new RGBColor();
        rgb.red = Math.round(255 * (1 - c) * (1 - k));
        rgb.green = Math.round(255 * (1 - m) * (1 - k));
        rgb.blue = Math.round(255 * (1 - y) * (1 - k));
        return rgb;
    }
    return null;
}

// 保護カラー（変更しないカラー）かどうかを判定（RGB変換後）
function isProtectedColor(color, includeBlack, excludeWhite) {
    if (!color || color.typename !== "RGBColor") return true;
    if (includeBlack) return false;
    if (color.red === 0 && color.green === 0 && color.blue === 0) return true;
    if (excludeWhite && color.red === 255 && color.green === 255 && color.blue === 255) return true;
    if (color.red === 255 || color.green === 255 || color.blue === 255) return true;
    if (color.red <= 51 && color.green <= 51 && color.blue <= 51) return true;
    return false;
}

// RGBColorの複製
function copyRGBColor(original) {
    var newColor = new RGBColor();
    newColor.red = original.red;
    newColor.green = original.green;
    newColor.blue = original.blue;
    return newColor;
}

// カラー一致をチェック
function colorsEqual(c1, c2) {
    return c1.red === c2.red && c1.green === c2.green && c1.blue === c2.blue;
}