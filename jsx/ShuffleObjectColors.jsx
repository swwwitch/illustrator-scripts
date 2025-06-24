#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**************************************************
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

【バグ】
- CMYK カラーの変換精度に注意（特に黒・白の扱い）
- 白、黒の保持の挙動に注意（除外条件の設定）

作成日：2025-06-24
更新日：2025-06-25
更新履歴：
- v1.0 初版リリース
- v1.1 バグフィックス
- v1.2 ランダム適用・順番適用の切替対応
**************************************************/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/*** 日英ラベル定義 Define label（UI出現順に整理） ***/
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "カラーをランダムに再適用", en: "Reapply Colors Randomly" },
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

// グローバル変数としてUIパーツを宣言
var fillCheckbox, strokeCheckbox, textCheckbox;
var blackCheckbox, whiteCheckbox;
var balanceCheckbox; // ←追加
var randomCheckbox; // ランダムチェックボックス

main();

// 塗り・線のカラー適用を共通化
function applyColorIfEligible(target, isFill, changeFlag, colorArray, colorIndex, excludeBlack, excludeWhite) {
    if (!changeFlag) return { color: null, index: colorIndex };

    var currentColor = isFill ? target.fillColor : target.strokeColor;
    var unified = getOriginalColor(currentColor);
    if (!unified || (excludeBlack && isBlack(unified)) || (excludeWhite && isWhite(unified))) return { color: null, index: colorIndex };

    var nextColorIndex = colorIndex % colorArray.length;
    var nextColor = getOriginalColor(colorArray[nextColorIndex]);

    if (isFill) {
        target.filled = true;
        if (target.typename === "TextFrame") {
            target.textRange.characterAttributes.fillColor = nextColor;
        } else {
            target.fillColor = nextColor;
        }
        // --- 文字ごとランダム適用 ---
        if (target.typename === "TextFrame" && typeof changeText !== "undefined" && changeText) {
            var chars = target.textRange.characters;
            for (var c = 0; c < chars.length; c++) {
                var color = getOriginalColor(colorArray[colorIndex % colorArray.length]);
                chars[c].characterAttributes.fillColor = color;
                colorIndex++;
            }
        }
    } else {
        target.stroked = true;
        if (target.typename === "TextFrame") {
            target.textRange.characterAttributes.strokeColor = nextColor;
        } else {
            target.strokeColor = nextColor;
        }
        // --- 文字ごとランダム適用 ---
        if (target.typename === "TextFrame" && typeof changeText !== "undefined" && changeText) {
            var chars = target.textRange.characters;
            for (var c = 0; c < chars.length; c++) {
                var color = getOriginalColor(colorArray[colorIndex % colorArray.length]);
                chars[c].characterAttributes.strokeColor = color;
                colorIndex++;
            }
        }
    }
    colorIndex++;
    return { color: nextColor, index: colorIndex };
}

/***********
メイン処理
***********/
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

    var textGroup = optionPanel.add("group");
    textGroup.orientation = "row";
    textGroup.alignChildren = "left";
    textCheckbox = textGroup.add("checkbox", undefined, "テキスト");
    textCheckbox.value = true;

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

    // 「ランダム」チェックボックス（パネル外・左下に追加）
    randomCheckbox = leftGroup.add("checkbox", undefined, LABELS.random[lang]);
    randomCheckbox.value = true;

    // 「バランスを保持」チェックボックス（パネル外・左下に追加）
    balanceCheckbox = leftGroup.add("checkbox", undefined, LABELS.balance[lang]);
    balanceCheckbox.value = false;

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
        dialog.close(0);
    };
    okBtn.onClick = function() {
        applyColors();
        dialog.close(1);
    };

    if (dialog.show() !== 1) return;
}

function applyColors() {
    var changeFill = fillCheckbox.value;
    var changeStroke = strokeCheckbox.value;
    var changeText = textCheckbox.value;
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
    collectAllTargetItems(selection, targetItems, changeText);

    var colorCountMap = {};
    var uniqueColorKeys = {};

    // カラー収集処理（共通）
    for (var i = 0; i < targetItems.length; i++) {
        var item = targetItems[i];

        if (changeFill && (item.filled || item.typename === "TextFrame")) {
            var fillColor = getOriginalColor(item.fillColor);
            if (
                fillColor &&
                !(excludeBlack && isBlack(fillColor)) &&
                !(excludeWhite && isWhite(fillColor))
            ) {
                collectColor(fillColor, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }
        if (changeFill && item.typename === "TextFrame" && changeText) {
            var chars = item.textRange.characters;
            for (var c = 0; c < chars.length; c++) {
                var charColor = getOriginalColor(chars[c].characterAttributes.fillColor);
                if (
                    charColor &&
                    !(excludeBlack && isBlack(charColor)) &&
                    !(excludeWhite && isWhite(charColor))
                ) {
                    collectColor(charColor, colorCountMap, uniqueColorKeys, balancePreserve);
                }
            }
        }

        if (changeStroke && (item.stroked || item.typename === "TextFrame") && item.strokeWidth >= 1) {
            var strokeColor = getOriginalColor(item.strokeColor);
            if (
                strokeColor &&
                !(excludeBlack && isBlack(strokeColor)) &&
                !(excludeWhite && isWhite(strokeColor))
            ) {
                collectColor(strokeColor, colorCountMap, uniqueColorKeys, balancePreserve);
            }
        }
        if (changeStroke && item.typename === "TextFrame" && changeText) {
            var chars = item.textRange.characters;
            for (var c = 0; c < chars.length; c++) {
                var charStroke = getOriginalColor(chars[c].characterAttributes.strokeColor);
                if (
                    charStroke &&
                    !(excludeBlack && isBlack(charStroke)) &&
                    !(excludeWhite && isWhite(charStroke))
                ) {
                    collectColor(charStroke, colorCountMap, uniqueColorKeys, balancePreserve);
                }
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

// カラー収集処理（共通）
function collectColor(color, colorCountMap, uniqueColorKeys, balancePreserve) {
    var key;
    if (color.typename === "RGBColor") {
        key = color.red + "," + color.green + "," + color.blue;
    } else if (color.typename === "CMYKColor") {
        key = color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black;
    } else {
        return;
    }
    if (balancePreserve) {
        if (!colorCountMap[key]) {
            colorCountMap[key] = { color: getOriginalColor(color), count: 0 };
        }
        colorCountMap[key].count++;
    } else if (!uniqueColorKeys[key]) {
        uniqueColorKeys[key] = getOriginalColor(color);
    }
}

// 再帰的に PathItem, TextFrame を収集
function collectAllTargetItems(items, results, includeText) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "GroupItem") {
            collectAllTargetItems(item.pageItems, results, includeText);
        } else if (item.typename === "CompoundPathItem") {
            collectAllTargetItems(item.pathItems, results, includeText);
        } else if (item.typename === "PathItem" || (item.typename === "TextFrame" && includeText)) {
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

// 任意のカラーをそのままコピー（RGB/CMYK対応）
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

function isBlack(color) {
    if (color.typename === "RGBColor") {
        return color.red <= 51 && color.green <= 51 && color.blue <= 51;
    } else if (color.typename === "CMYKColor") {
        // 黒に近いか判定（CMYKで黒判定は厳密には難しいが単純判定）
        return color.black >= 0.8 || (color.cyan <= 0.2 && color.magenta <= 0.2 && color.yellow <= 0.2 && color.black >= 0.7);
    }
    return false;
}

function isWhite(color) {
    if (color.typename === "RGBColor") {
        return color.red === 255 && color.green === 255 && color.blue === 255;
    } else if (color.typename === "CMYKColor") {
        // CMYKで白は全て0
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }
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