#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

ApplySwatchesToSelection.jsx

### 概要

- 選択したオブジェクトまたはテキストに、スウォッチや定義済みカラーを自動適用するスクリプトです。
- CMYK / RGB カラーモードに応じてカラーを使い分けます。

### 主な機能

- スウォッチパネル選択スウォッチ、またはプロセススウォッチの利用
- テキスト単体は文字単位、複数オブジェクトは位置順にカラーを適用
- 3色以上のスウォッチ時にはランダム適用
- 日本語／英語UI切替対応

### 処理の流れ

1. ドキュメントと選択確認
2. スウォッチ取得（未選択時は定義済みカラーを使用）
3. テキストは文字単位、オブジェクトは位置順にカラー適用
4. 必要に応じてランダムシャッフル

### 更新履歴

- v1.0.0 (20241103) : 初期バージョン
- v1.1.0 (20250625) : スウォッチ未選択時の全プロセス対応
- v1.2.0 (20250708) : CMYK/RGB切替対応

---

### Script Name:

ApplySwatchesToSelection.jsx

### Overview

- A script to automatically apply swatches or predefined colors to selected objects or text.
- Switches colors depending on document color mode (CMYK or RGB).

### Main Features

- Use selected swatches from panel or all process swatches
- Apply per character for single text, or per object in order
- Random shuffle when more than 3 swatches
- Supports Japanese / English UI

### Process Flow

1. Check document and selection
2. Get swatches (use predefined colors if none selected)
3. Apply per character or per object in order
4. Shuffle randomly if needed

### Update History

- v1.0.0 (20241103): Initial version
- v1.1.0 (20250625): Added process swatch fallback when no swatches selected
- v1.2.0 (20250708): Added CMYK/RGB mode switch support

*/

function getCurrentLang() {
    var locale = $.locale.toLowerCase();
    if (locale.indexOf('ja') === 0) {
        return 'ja';
    }
    return 'en';
}

// エラーメッセージなどのラベル（UI表示順に整理）
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
        // ドキュメントが開かれているか確認
        if (app.documents.length === 0) {
            alert(LABELS.errNoDoc[lang]);
            return;
        }
        var activeDoc = app.activeDocument;
        var selectedItems = app.selection;

        // 選択オブジェクトがあるか確認
        if (selectedItems.length === 0) {
            alert(LABELS.errNoSelection[lang]);
            return;
        }

        // 選択されたスウォッチを取得
        var selectedSwatches = activeDoc.swatches.getSelected();

        // 定義済みのカラーセット（CMYK / RGB）
        var predefinedColors = [];
        if (activeDoc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk1 = new CMYKColor();
            cmyk1.cyan = 9; cmyk1.magenta = 80; cmyk1.yellow = 95; cmyk1.black = 0;
            predefinedColors.push(cmyk1);
            var cmyk2 = new CMYKColor();
            cmyk2.cyan = 7; cmyk2.magenta = 3; cmyk2.yellow = 86; cmyk2.black = 0;
            predefinedColors.push(cmyk2);
            var cmyk3 = new CMYKColor();
            cmyk3.cyan = 76; cmyk3.magenta = 8; cmyk3.yellow = 100; cmyk3.black = 0;
            predefinedColors.push(cmyk3);
            var cmyk4 = new CMYKColor();
            cmyk4.cyan = 72; cmyk4.magenta = 22; cmyk4.yellow = 6; cmyk4.black = 0;
            predefinedColors.push(cmyk4);
            var cmyk5 = new CMYKColor();
            cmyk5.cyan = 38; cmyk5.magenta = 54; cmyk5.yellow = 79; cmyk5.black = 0;
            predefinedColors.push(cmyk5);
            var cmyk6 = new CMYKColor();
            cmyk6.cyan = 6; cmyk6.magenta = 36; cmyk6.yellow = 84; cmyk6.black = 0;
            predefinedColors.push(cmyk6);
        } else {
            var rgb1 = new RGBColor();
            rgb1.red = 222; rgb1.green = 84; rgb1.blue = 25;
            predefinedColors.push(rgb1);
            var rgb2 = new RGBColor();
            rgb2.red = 245; rgb2.green = 233; rgb2.blue = 40;
            predefinedColors.push(rgb2);
            var rgb3 = new RGBColor();
            rgb3.red = 41; rgb3.green = 163; rgb3.blue = 57;
            predefinedColors.push(rgb3);
            var rgb4 = new RGBColor();
            rgb4.red = 53; rgb4.green = 157; rgb4.blue = 209;
            predefinedColors.push(rgb4);
            var rgb5 = new RGBColor();
            rgb5.red = 173; rgb5.green = 127; rgb5.blue = 71;
            predefinedColors.push(rgb5);
            var rgb6 = new RGBColor();
            rgb6.red = 238; rgb6.green = 176; rgb6.blue = 51;
            predefinedColors.push(rgb6);
        }

        // スウォッチが未選択、または1色以下、または白のみの場合は定義済みカラーを使用
        if (!selectedSwatches || selectedSwatches.length <= 1 || allWhiteSwatches(selectedSwatches)) {
            selectedSwatches = [];
            for (var i = 0; i < predefinedColors.length; i++) {
                var dummySwatch = {};
                dummySwatch.color = predefinedColors[i];
                // 削除: dummySwatch.opacity = 100; // 不透明度100%を明示的に設定
                selectedSwatches.push(dummySwatch);
            }
        }

        // スウォッチ数が3色より多い場合はランダムにシャッフル
        if (selectedSwatches.length > 3) {
            selectedSwatches = shuffleArray(selectedSwatches);
        }

        // 選択が単一テキストフレームの場合は文字単位で色付け
        if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
            var selectedTextFrame = selectedItems[0];
            var charCount = selectedTextFrame.contents.length;
            for (var i = 0; i < charCount; i++) {
                var swatchColor = getSwatchColor(i, selectedSwatches);
                selectedTextFrame.characters[i].fillColor = swatchColor;
                selectedTextFrame.characters[i].strokeColor = new NoColor();
                selectedTextFrame.characters[i].opacity = 100;
            }
        } else {
            // 複数オブジェクトは位置順に並べ替えて色付け
            sortByPosition(selectedItems);
            for (var i = 0; i < selectedItems.length; i++) {
                var swatchColor = getSwatchColor(i, selectedSwatches);
                var currentItem = selectedItems[i];
                if (currentItem.typename === "PathItem") {
                    currentItem.fillColor = swatchColor;
                    currentItem.stroked = false;
                    currentItem.opacity = 100;
                } else if (currentItem.typename === "CompoundPathItem" && currentItem.pathItems.length > 0) {
                    var pathItems = currentItem.pathItems;
                    for (var j = 0; j < pathItems.length; j++) {
                        pathItems[j].fillColor = swatchColor;
                        pathItems[j].stroked = false;
                        pathItems[j].opacity = 100;
                    }
                } else if (currentItem.typename === "TextFrame") {
                    currentItem.textRange.fillColor = swatchColor;
                    currentItem.textRange.strokeColor = new NoColor();
                    currentItem.textRange.opacity = 100;
                }
            }
        }
    } catch (e) {
        alert(LABELS.errUnexpected[lang] + e.message);
    }
}

// 指定ドキュメントから使用可能なプロセススウォッチを取得
function getAvailableProcessSwatches(doc) {
    var result = [];
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        var col = swatches[i].color;
        // スポットカラー、グラデーション、パターン、グレースケール以外で、登録色でなく、白色でないもの
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

// 色が白かどうか判定（CMYK=0,0,0,0 または RGB=255,255,255）
function isWhiteColor(color) {
    if (color.typename === "CMYKColor") {
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    } else if (color.typename === "RGBColor") {
        return color.red === 255 && color.green === 255 && color.blue === 255;
    }
    return false;
}

// オブジェクト配列を位置順にソート（横幅が広ければ左→右、縦幅が広ければ上→下）
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
        // 横幅が広い場合は左から右へ
        items.sort(function(a, b) { return compPosition(a.left, b.left, b.top, a.top); });
    } else {
        // 縦幅が広い場合は上から下へ
        items.sort(function(a, b) { return compPosition(b.top, a.top, a.left, b.left); });
    }
}

// ソート用比較関数（主キー比較、同値なら副キー比較）
function compPosition(a1, b1, a2, b2) {
    return a1 == b1 ? a2 - b2 : a1 - b1;
}

// 配列をランダムシャッフルして返す
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

// インデックスに応じてスウォッチの色を取得（ループ）
function getSwatchColor(index, swatches) {
    var swatch = swatches[index % swatches.length];
    var color = swatch.color;
    // 定義済みカラーの場合、常に100%不透明度を返す
    if (typeof swatch.opacity !== "undefined") {
        color.opacity = swatch.opacity;
    } else {
        color.opacity = 100;
    }
    return color;
}

// 全スウォッチが白色のみか判定
function allWhiteSwatches(swatches) {
    for (var i = 0; i < swatches.length; i++) {
        if (!isWhiteColor(swatches[i].color)) {
            return false;
        }
    }
    return true;
}

main();