#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトやテキストに、スウォッチや定義済みカラーを自動で適用します。
CMYK／RGBのカラーモードに応じて使う色を切り替え、CMYKでスウォッチ未選択のときは2色混合の色をランダム生成します。

詳細は README を参照してください。

### Overview

Applies swatches, or predefined colors, to the selected objects and text automatically.
The palette follows the CMYK/RGB color mode, and in CMYK with no swatch selected it generates random two-ink combinations.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplySwatchesToSelection";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplySwatchesToSelection.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplySwatchesToSelection.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// CMYK fallback generation: maximum total (C+M or C+Y or M+Y)
var TMK_CMYK_FALLBACK_MAX_TOTAL = 200;

// CMYK fallback generation: minimum distance (Manhattan) between generated colors to avoid similar colors
var TMK_CMYK_FALLBACK_MIN_DISTANCE = 35;

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
        // 選択をフラット化（グループ内のテキスト/パスも対象にする）
        var selectedItems = flattenSelection(app.selection);
        // テキスト編集中に文字範囲が選択されている場合（TextRange）
        var selectedTextRange = getSingleSelectedTextRange(app.selection);

        // 選択オブジェクトがあるか確認
        if ((selectedItems.length === 0) && !selectedTextRange) {
            alert(LABELS.errNoSelection[lang]);
            return;
        }

        // 選択されたスウォッチを取得
        var selectedSwatches = activeDoc.swatches.getSelected();

        // 定義済みのカラーセット（CMYK / RGB）
        var predefinedColors = [];
        if (activeDoc.documentColorSpace === DocumentColorSpace.CMYK) {
            // CMYK fallback palette will be generated on-demand (depends on target count)
            predefinedColors = [];
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
            // 選択オブジェクトがちょうど4つの場合、固定の4色プリセットを適用
            if (selectedItems.length === 4 && !selectedTextRange) {
                predefinedColors = getFourColorPreset(activeDoc.documentColorSpace);
            } else if (activeDoc.documentColorSpace === DocumentColorSpace.CMYK) {
                // If CMYK document, generate as many colors as targets (avoid duplicates when possible)
                var needCount = getNeededColorCount(selectedItems, selectedTextRange);
                predefinedColors = generateRandomCMYPaletteUnique(needCount, TMK_CMYK_FALLBACK_MAX_TOTAL);
            }
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

        // 単一テキスト（テキストフレーム or TextRange）の場合は文字単位で色付け
        if (selectedTextRange) {
            var chars = selectedTextRange.characters;
            for (var i = 0; i < chars.length; i++) {
                var swatchColor = getSwatchColor(i, selectedSwatches);
                chars[i].fillColor = swatchColor;
                chars[i].strokeColor = new NoColor();
                chars[i].opacity = 100;
            }
        } else if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
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
                    // 複数テキストはテキストオブジェクト単位で色付け
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

// 選択をフラット化して、色付け対象（PathItem / CompoundPathItem / TextFrame）だけを収集
function flattenSelection(selection) {
    var result = [];
    if (!selection || selection.length === 0) return result;

    for (var i = 0; i < selection.length; i++) {
        collectColorTargets(selection[i], result);
    }

    return result;
}

function collectColorTargets(item, outArr) {
    if (!item) return;

    // 直接対象
    if (item.typename === "PathItem" || item.typename === "CompoundPathItem" || item.typename === "TextFrame") {
        outArr.push(item);
        return;
    }

    // グループ内を再帰的に探索
    if (item.typename === "GroupItem") {
        var pageItems = item.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            collectColorTargets(pageItems[i], outArr);
        }
        return;
    }

    // それ以外（PlacedItem 等）は無視
}

// テキスト編集中に文字範囲が選択されているケース（TextRange）を取得
function getSingleSelectedTextRange(selection) {
    try {
        if (!selection || selection.length !== 1) return null;
        if (selection[0] && selection[0].typename === "TextRange") return selection[0];
    } catch (e) { }
    return null;
}

// 処理対象数（必要な色数）を算出
function getNeededColorCount(selectedItems, selectedTextRange) {
    try {
        if (selectedTextRange) {
            return Math.max(1, selectedTextRange.characters.length);
        }
        if (selectedItems && selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
            return Math.max(1, selectedItems[0].contents.length);
        }
        if (selectedItems && selectedItems.length > 0) {
            return Math.max(1, selectedItems.length);
        }
    } catch (e) { }
    return 1;
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
        items.sort(function (a, b) { return compPosition(a.left, b.left, b.top, a.top); });
    } else {
        // 縦幅が広い場合は上から下へ
        items.sort(function (a, b) { return compPosition(b.top, a.top, a.left, b.left); });
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

// 乱数（整数）
function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

// CMYの距離（マンハッタン距離）
function cmyDistance(c1, m1, y1, c2, m2, y2) {
    return Math.abs(c1 - c2) + Math.abs(m1 - m2) + Math.abs(y1 - y2);
}

// 既存候補と十分離れているか
function isFarEnoughCMY(c, m, y, existing, minDist) {
    for (var i = 0; i < existing.length; i++) {
        var e = existing[i];
        if (cmyDistance(c, m, y, e.cyan, e.magenta, e.yellow) < minDist) {
            return false;
        }
    }
    return true;
}

// CMYKドキュメント用：CM/CY/MY の2チャンネルのみ（K=0）で、合計が maxTotal を超えない範囲の色を生成
function generateRandomCMYPalette(count, maxTotal) {
    var result = [];
    var pairs = ["CM", "CY", "MY"];

    for (var i = 0; i < count; i++) {
        var pair = pairs[i % pairs.length];
        var c = 0, m = 0, y = 0;

        // Generate two non-zero channels with a+b <= maxTotal
        var guard = 0;
        while (guard++ < 200) {
            // Keep values in 1..100, but enforce total <= maxTotal
            var a = randInt(1, Math.min(100, maxTotal - 1));
            var bMax = Math.min(100, maxTotal - a);
            if (bMax < 1) continue;
            var b = randInt(1, bMax);

            if (pair === "CM") { c = a; m = b; y = 0; }
            else if (pair === "CY") { c = a; y = b; m = 0; }
            else { m = a; y = b; c = 0; }
            break;
        }

        var col = new CMYKColor();
        col.cyan = c;
        col.magenta = m;
        col.yellow = y;
        col.black = 0;
        result.push(col);
    }

    return result;
}

// CMYKドキュメント用：CM/CY/MY の2チャンネルのみ（K=0）で、可能な限り重複しない色を生成
// ※対象数が非常に多い場合、理論上の組み合わせ上限に達すると重複を許容する
function generateRandomCMYPaletteUnique(count, maxTotal) {
    var result = [];
    var accepted = []; // 既に採用した色（近い色回避用）
    var seen = {};     // 完全一致回避
    var pairs = ["CM", "CY", "MY"];

    var minDist = TMK_CMYK_FALLBACK_MIN_DISTANCE;
    var maxTries = Math.max(1500, count * 80);
    var tries = 0;

    while (result.length < count && tries++ < maxTries) {
        // 生成が詰まるときは距離制約を徐々に緩める
        if (minDist > 0 && (tries % 500) === 0) {
            minDist = Math.max(0, minDist - 5);
        }

        // ペア選択：パターン固定を避けるため、たまに並びをシャッフル
        var pair = pairs[result.length % pairs.length];
        if ((tries % 37) === 0) {
            pairs = shuffleArray(pairs);
            pair = pairs[result.length % pairs.length];
        }

        var c = 0, m = 0, y = 0;

        // 2チャンネル非ゼロ、合計 <= maxTotal
        var a = randInt(1, Math.min(100, maxTotal - 1));
        var bMax = Math.min(100, maxTotal - a);
        if (bMax < 1) continue;
        var b = randInt(1, bMax);

        if (pair === "CM") { c = a; m = b; y = 0; }
        else if (pair === "CY") { c = a; y = b; m = 0; }
        else { m = a; y = b; c = 0; }

        var key = c + "," + m + "," + y;
        if (seen[key]) continue;

        // 近い色を回避
        if (minDist > 0 && !isFarEnoughCMY(c, m, y, accepted, minDist)) {
            continue;
        }

        seen[key] = true;

        var col = new CMYKColor();
        col.cyan = c;
        col.magenta = m;
        col.yellow = y;
        col.black = 0;

        result.push(col);
        accepted.push(col);
    }

    // どうしても埋まらない場合：重複許容。ただし距離制約はベストエフォート
    var guard = 0;
    while (result.length < count) {
        if (guard++ > 3000) {
            minDist = 0; // 完走優先
        }

        var pair2 = pairs[result.length % pairs.length];
        if ((guard % 37) === 0) {
            pairs = shuffleArray(pairs);
            pair2 = pairs[result.length % pairs.length];
        }

        var c2 = 0, m2 = 0, y2 = 0;

        var a2 = randInt(1, Math.min(100, maxTotal - 1));
        var bMax2 = Math.min(100, maxTotal - a2);
        if (bMax2 < 1) continue;
        var b2 = randInt(1, bMax2);

        if (pair2 === "CM") { c2 = a2; m2 = b2; y2 = 0; }
        else if (pair2 === "CY") { c2 = a2; y2 = b2; m2 = 0; }
        else { m2 = a2; y2 = b2; c2 = 0; }

        if (minDist > 0 && !isFarEnoughCMY(c2, m2, y2, accepted, minDist)) {
            continue;
        }

        var col2 = new CMYKColor();
        col2.cyan = c2;
        col2.magenta = m2;
        col2.yellow = y2;
        col2.black = 0;

        result.push(col2);
        accepted.push(col2);
    }

    return result;
}

// 4オブジェクト選択時の固定プリセット (#B9D3E0, #E19DA1, #FDECAC, #CB4447)
function getFourColorPreset(colorSpace) {
    var colors = [];
    if (colorSpace === DocumentColorSpace.CMYK) {
        var defs = [
            { c: 19, m: 7, y: 2, k: 12 },   // #B9D3E0
            { c: 0, m: 30, y: 28, k: 12 },   // #E19DA1
            { c: 0, m: 7, y: 32, k: 1 },     // #FDECAC
            { c: 0, m: 67, y: 65, k: 20 }    // #CB4447
        ];
        for (var i = 0; i < defs.length; i++) {
            var col = new CMYKColor();
            col.cyan = defs[i].c;
            col.magenta = defs[i].m;
            col.yellow = defs[i].y;
            col.black = defs[i].k;
            colors.push(col);
        }
    } else {
        var defs = [
            { r: 185, g: 211, b: 224 },  // #B9D3E0
            { r: 225, g: 157, b: 161 },  // #E19DA1
            { r: 253, g: 236, b: 172 },  // #FDECAC
            { r: 203, g: 68, b: 71 }     // #CB4447
        ];
        for (var i = 0; i < defs.length; i++) {
            var col = new RGBColor();
            col.red = defs[i].r;
            col.green = defs[i].g;
            col.blue = defs[i].b;
            colors.push(col);
        }
    }
    return colors;
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
