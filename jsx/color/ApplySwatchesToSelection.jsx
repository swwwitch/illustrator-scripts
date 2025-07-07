#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ApplySwatchesToSelection.jsx

### 概要

- 選択中のオブジェクトやテキストに、スウォッチパネルのスウォッチまたは全プロセススウォッチからカラーを適用するスクリプトです。
- 適用方法は順番またはランダムから選択でき、文字単位やオブジェクト単位でカラーを変えることができます。

### 主な機能

- スウォッチパネルの選択スウォッチ、または全プロセススウォッチを使用
- 単一テキストの場合は文字ごとに色分け
- 複数オブジェクトの場合は位置順にカラー適用
- 3色以上ではスウォッチをランダムシャッフル
- 日本語／英語インターフェース対応

### 処理の流れ

1. ドキュメントと選択オブジェクトの確認
2. スウォッチ取得（選択されていなければ全プロセススウォッチ）
3. 単一テキストは文字単位、複数オブジェクトは位置順に色を適用
4. 必要に応じてランダムシャッフル

### オリジナル、謝辞

sort_by_position.jsx（shspage氏）を参考にしました  
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

### 更新履歴

- v1.0.0 (20241103) : 初期バージョン
- v1.1.0 (20250625) : スウォッチ未選択時に全プロセススウォッチ対応
- v1.2.0 (20250708) : 微調整

---

### Script Name:

ApplySwatchesToSelection.jsx

### Overview

- A script that applies colors from selected swatches or all process swatches to selected objects or text.
- You can choose to apply colors sequentially or randomly, and assign colors per character or per object.

### Main Features

- Uses selected swatches from the Swatches panel or all process swatches
- Colors individual characters for single text frame
- Colors objects in order when multiple objects are selected
- Random shuffle if there are more than 3 colors
- Japanese and English UI support

### Process Flow

1. Check document and selection
2. Get swatches (if none selected, use all process swatches)
3. Apply colors per character for single text, or in order for multiple objects
4. Shuffle randomly if needed

### Original / Acknowledgements

Inspired by sort_by_position.jsx by shspage  
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

### Update History

- v1.0.0 (20241103): Initial version
- v1.1.0 (20250625): Supported all process swatches when no swatches are selected
*/

function getCurrentLang() {
    var locale = $.locale.toLowerCase();
    if (locale.indexOf('ja') === 0) {
        return 'ja';
    }
    return 'en';
}

var LABELS = {
    errNoDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errNoSelection: {
        ja: "オブジェクトを選択してください。",
        en: "Please select objects."
    },
    errUnexpected: {
        ja: "エラーが発生しました: ",
        en: "An error occurred: "
    }
};

function main() {
    var lang = getCurrentLang();
    try {
        if (app.documents.length === 0) {
            alert(LABELS.errNoDoc[lang]);
            return;
        }
        var activeDoc = app.activeDocument;

        var selectedItems = app.selection;

        if (selectedItems.length > 0) {
            var selectedSwatches = activeDoc.swatches.getSelected();

            // スウォッチが選択されていない、または白のみ選択の場合はプロセスカラースウォッチをすべて取得し、ランダムに利用
            if (!selectedSwatches || selectedSwatches.length === 0 || allWhiteSwatches(selectedSwatches)) {
                selectedSwatches = getAvailableProcessSwatches(activeDoc);
            }

            // スウォッチ数が3色より多い場合はランダムにシャッフル
            if (selectedSwatches.length > 3) {
                selectedSwatches = shuffleArray(selectedSwatches);
            }

            if (selectedSwatches.length > 0) {
                if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
                    var selectedTextFrame = selectedItems[0];
                    var charCount = selectedTextFrame.contents.length;
                    for (var i = 0; i < charCount; i++) {
                        var swatchColor = getSwatchColor(i, selectedSwatches);
                        selectedTextFrame.characters[i].fillColor = swatchColor;
                    }
                } else {
                    // 複数オブジェクトは位置順に並べ替えてスウォッチを適用
                    sortByPosition(selectedItems);
                    for (var i = 0; i < selectedItems.length; i++) {
                        var swatchColor = getSwatchColor(i, selectedSwatches);
                        var currentItem = selectedItems[i];
                        if (currentItem.typename === "PathItem") {
                            currentItem.fillColor = swatchColor;
                        } else if (currentItem.typename === "CompoundPathItem" && currentItem.pathItems.length > 0) {
                            var pathItems = currentItem.pathItems;
                            for (var j = 0; j < pathItems.length; j++) {
                                pathItems[j].fillColor = swatchColor;
                            }
                        } else if (currentItem.typename === "TextFrame") {
                            currentItem.textRange.fillColor = swatchColor;
                        }
                    }
                }
            } else {
                // ここは実質発生しないが念のため
                alert(LABELS.errNoSelection[lang]);
            }
        } else {
            alert(LABELS.errNoSelection[lang]);
        }
    } catch (e) {
        alert(LABELS.errUnexpected[lang] + e.message);
    }
}

// 使用可能なスウォッチ（プロセスカラー）だけを取得
function getAvailableProcessSwatches(doc) {
    var result = [];
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        var col = swatches[i].color;
        if (
            !(col.typename === "SpotColor" || col.typename === "GradientColor" || col.typename === "PatternColor" || col.typename === "GrayColor")
            && swatches[i].name !== "[Registration]"
            && !isWhiteColor(col)
        ) {
            result.push(swatches[i]);
        }
    }
    return result;
}

// 白色(CMYK=0,0,0,0またはRGB=255,255,255)かどうか判定
function isWhiteColor(color) {
    if (color.typename === "CMYKColor") {
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    } else if (color.typename === "RGBColor") {
        return color.red === 255 && color.green === 255 && color.blue === 255;
    }
    return false;
}

// オブジェクトを位置順に並べ替える（横幅が広ければ左→右、縦幅が広ければ上→下）
function sortByPosition(items) {
    var hMin = Infinity, hMax = -Infinity, vMin = Infinity, vMax = -Infinity;
    for (var i = 0, len = items.length; i < len; i++) {
        var left = items[i].left;
        var top = items[i].top;
        if (left < hMin) hMin = left;
        if (left > hMax) hMax = left;
        if (top < vMin) vMin = top;
        if (top > vMax) vMax = top;
    }
    if (hMax - hMin > vMax - vMin) {
        items.sort(function(a, b) { return compPosition(a.left, b.left, b.top, a.top); });
    } else {
        items.sort(function(a, b) { return compPosition(b.top, a.top, a.left, b.left); });
    }
}

// 並べ替え用の比較関数
function compPosition(a1, b1, a2, b2) {
    return a1 == b1 ? a2 - b2 : a1 - b1;
}

// 配列をシャッフルする処理
function shuffleArray(arr) {
    var result = arr.slice();
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return result;
}

// インデックスに応じてスウォッチの色を取得
function getSwatchColor(index, swatches) {
    return swatches[index % swatches.length].color;
}

main();

// 白のみ選択されている場合に true を返す
function allWhiteSwatches(swatches) {
    for (var i = 0; i < swatches.length; i++) {
        if (!isWhiteColor(swatches[i].color)) {
            return false;
        }
    }
    return true;
}