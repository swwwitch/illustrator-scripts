#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：ApplySwatchesToSelection.jsx

概要：
スウォッチパネルで選択されているスウォッチ、または全スウォッチの中からプロセスカラーを使い、
選択オブジェクトやテキストに順またはランダムでカラーを適用します。

処理の流れ：
1. ドキュメントと選択状態の確認
2. スウォッチの取得（未選択時はプロセスカラーからランダム取得）
3. テキスト1つ選択時は文字ごとにスウォッチを適用
4. 複数オブジェクト選択時は位置順に並べ替えてスウォッチを適用

対象：TextFrame, PathItem, CompoundPathItem（内部のPathItem含む）
限定条件：オブジェクトが選択されていること

謝辞：
sort_by_position.jsx（shspage氏）を参考にしました。
https://gist.github.com/shspage/02c6d8654cf6b3798b6c0b69d976a891

作成日：2024年11月03日
最終更新日：2025年06月25日
- v1.1 スウォッチを選択していないときには、全スウォッチを対象に
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

            // スウォッチが選択されていない場合はプロセスカラースウォッチをすべて取得し、ランダムに利用
            if (selectedSwatches.length === 0) {
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
        if (!(col.typename === "SpotColor" || col.typename === "GradientColor" || col.typename === "PatternColor")) {
            result.push(swatches[i]);
        }
    }
    return result;
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