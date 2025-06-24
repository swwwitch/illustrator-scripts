#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：ApplySwatchesToSelection.jsx
作成日：2024年11月03日
更新日：2025年06月24日

概要：
選択中のオブジェクトに対して、スウォッチパネルで選択されているスウォッチを順に適用します。

処理の流れ：
1. ドキュメントと選択状態の確認
2. スウォッチの選択を取得し、必要に応じてランダム化
3. テキスト1つ選択時は文字ごとにスウォッチを適用
4. 複数オブジェクト選択時は位置順に並べ替えてスウォッチを適用

その他
- スウォッチ数が3色より多い場合には、ランダムに適用される（var randomizeSwatches = true;）

対象と対象外：
対象：TextFrame, PathItem, CompoundPathItem（内部のPathItemも含む）
対象外：上記以外のオブジェクト

限定条件：
スウォッチが1つ以上選択されていること、オブジェクトが選択されていること

謝辞：
縦方向の調整にはshspageさんのsort_by_position.jsxを参照しました。
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

解説記事：
https://note.com/dtp_tranist/n/n51f21e1526f9

creoldさんが多機能版を公開されています。
https://github.com/creold/illustrator-scripts/blob/master/md/Color.md#matchcolors

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
    errNoSwatch: {
        ja: "スウォッチパネルでスウォッチを選択してください。",
        en: "Please select swatches in the Swatches panel."
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

            // スウォッチ数が3色より多い場合のみランダム化
            var randomizeSwatches = true;
            if (randomizeSwatches && selectedSwatches.length > 3) {
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
                    // 位置順に並べ替えてスウォッチを適用
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
                alert(LABELS.errNoSwatch[lang]);
            }
        } else {
            alert(LABELS.errNoSelection[lang]);
        }
    } catch (e) {
        alert(LABELS.errUnexpected[lang] + e.message);
    }
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