#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトやテキストに、スウォッチや定義済みカラーを自動で適用します。
CMYK／RGBのカラーモードに応じて使う色を切り替え、CMYKでスウォッチ未選択のときは2色混合の色をランダム生成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplySwatchesToSelection.md

### Overview

Applies swatches, or predefined colors, to the selected objects and text automatically.
The palette follows the CMYK/RGB color mode, and in CMYK with no swatch selected it generates random two-ink combinations.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplySwatchesToSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplySwatchesToSelection";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplySwatchesToSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplySwatchesToSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* RGB ドキュメントでスウォッチ未選択のときに使う色 / Colors used in RGB documents when no swatches are selected */
    var RGB_FALLBACK_COLORS = [
        [222, 84, 25],
        [245, 233, 40],
        [41, 163, 57],
        [53, 157, 209],
        [173, 127, 71],
        [238, 176, 51]
    ];

    /* ちょうど4つ選択したときの固定プリセット (#B9D3E0, #E19DA1, #FDECAC, #CB4447) / Fixed preset for exactly four objects */
    var FOUR_COLOR_PRESET_CMYK = [
        [19, 7, 2, 12],   /* #B9D3E0 */
        [0, 30, 28, 12],  /* #E19DA1 */
        [0, 7, 32, 1],    /* #FDECAC */
        [0, 67, 65, 20]   /* #CB4447 */
    ];
    var FOUR_COLOR_PRESET_RGB = [
        [185, 211, 224],  /* #B9D3E0 */
        [225, 157, 161],  /* #E19DA1 */
        [253, 236, 172],  /* #FDECAC */
        [203, 68, 71]     /* #CB4447 */
    ];

    /* CMYK で自動生成する2色混合の合計の上限（C+M / C+Y / M+Y） / Max total of the two inks in generated CMYK colors */
    var CMYK_FALLBACK_MAX_TOTAL = 200;

    /* 自動生成する色どうしの最小距離（マンハッタン距離、似た色を避ける） / Min Manhattan distance between generated colors */
    var CMYK_FALLBACK_MIN_DISTANCE = 35;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            unexpected: { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // 選択の取得 / Selection
    // =========================================

    /**
     * 色付け対象（パス・複合パス・テキスト）を集める。グループ内は再帰的にたどる
     * @param {PageItem} item - 対象のオブジェクト
     * @param {PageItem[]} outItems - 集めたオブジェクトの格納先
     * @returns {void}
     */
    function collectColorTargets(item, outItems) {
        if (!item) return;
        if (item.typename === "PathItem" || item.typename === "CompoundPathItem" || item.typename === "TextFrame") {
            outItems.push(item);
            return;
        }
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectColorTargets(item.pageItems[i], outItems);
            }
        }
        /* それ以外（配置画像など）は無視 / Ignore other items such as placed images */
    }

    /**
     * 選択をフラットにして、色付け対象だけを集める
     * @param {Array} selectedItems - 選択
     * @returns {PageItem[]} 色付け対象
     */
    function flattenSelection(selectedItems) {
        var colorTargets = [];
        if (!selectedItems || selectedItems.length === 0) return colorTargets;
        for (var i = 0; i < selectedItems.length; i++) {
            collectColorTargets(selectedItems[i], colorTargets);
        }
        return colorTargets;
    }

    /**
     * テキスト編集中に文字範囲が選択されていれば、その TextRange を返す
     * @param {Array} selectedItems - 選択
     * @returns {TextRange|null} 文字範囲（なければ null）
     */
    function getSingleSelectedTextRange(selectedItems) {
        if (!selectedItems || selectedItems.length !== 1) return null;
        if (selectedItems[0] && selectedItems[0].typename === "TextRange") return selectedItems[0];
        return null;
    }

    /**
     * テキストフレーム1つだけが選択されているか
     * @param {PageItem[]} colorTargets - 色付け対象
     * @returns {boolean} 1つだけで TextFrame なら true
     */
    function isSingleTextFrame(colorTargets) {
        return colorTargets.length === 1 && colorTargets[0].typename === "TextFrame";
    }

    /**
     * 必要な色の数（文字数またはオブジェクト数）を返す
     * @param {PageItem[]} colorTargets - 色付け対象
     * @param {TextRange|null} selectedTextRange - 選択中の文字範囲
     * @returns {number} 色の数（1以上）
     */
    function getNeededColorCount(colorTargets, selectedTextRange) {
        if (selectedTextRange) {
            return Math.max(1, selectedTextRange.characters.length);
        }
        if (isSingleTextFrame(colorTargets)) {
            return Math.max(1, colorTargets[0].contents.length);
        }
        return Math.max(1, colorTargets.length);
    }

    // =========================================
    // カラー / Colors
    // =========================================

    /**
     * RGB カラーを作る
     * @param {number[]} rgbValues - [R, G, B]
     * @returns {RGBColor} カラー
     */
    function createRGBColor(rgbValues) {
        var color = new RGBColor();
        color.red = rgbValues[0];
        color.green = rgbValues[1];
        color.blue = rgbValues[2];
        return color;
    }

    /**
     * CMYK カラーを作る
     * @param {number[]} cmykValues - [C, M, Y, K]
     * @returns {CMYKColor} カラー
     */
    function createCMYKColor(cmykValues) {
        var color = new CMYKColor();
        color.cyan = cmykValues[0];
        color.magenta = cmykValues[1];
        color.yellow = cmykValues[2];
        color.black = cmykValues[3];
        return color;
    }

    /**
     * 値の表からカラーの配列を作る
     * @param {number[][]} valueTable - カラー値の配列
     * @param {Function} createColor - createRGBColor または createCMYKColor
     * @returns {Color[]} カラーの配列
     */
    function createColors(valueTable, createColor) {
        var colors = [];
        for (var i = 0; i < valueTable.length; i++) {
            colors.push(createColor(valueTable[i]));
        }
        return colors;
    }

    /**
     * 4つ選択したときの固定プリセットを返す
     * @param {DocumentColorSpace} colorSpace - ドキュメントのカラーモード
     * @returns {Color[]} 4色
     */
    function getFourColorPreset(colorSpace) {
        if (colorSpace === DocumentColorSpace.CMYK) {
            return createColors(FOUR_COLOR_PRESET_CMYK, createCMYKColor);
        }
        return createColors(FOUR_COLOR_PRESET_RGB, createRGBColor);
    }

    /**
     * 白（CMYK=0,0,0,0 または RGB=255,255,255）か判定する
     * @param {Color} color - 判定するカラー
     * @returns {boolean} 白なら true
     */
    function isWhiteColor(color) {
        if (color.typename === "CMYKColor") {
            return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        } else if (color.typename === "RGBColor") {
            return color.red === 255 && color.green === 255 && color.blue === 255;
        }
        return false;
    }

    /**
     * すべてのスウォッチが白か判定する
     * @param {Swatch[]} swatches - スウォッチ
     * @returns {boolean} すべて白なら true
     */
    function allWhiteSwatches(swatches) {
        for (var i = 0; i < swatches.length; i++) {
            if (!isWhiteColor(swatches[i].color)) {
                return false;
            }
        }
        return true;
    }

    /**
     * インデックスに応じてスウォッチの色を取り出す（数が足りなければ繰り返す）
     * @param {number} index - 何番目か
     * @param {Array<{color: Color}>} swatches - スウォッチ（または color を持つオブジェクト）
     * @returns {Color} カラー
     */
    function getSwatchColor(index, swatches) {
        var swatch = swatches[index % swatches.length];
        var color = swatch.color;
        /* 定義済みカラーは常に100% / Predefined colors are always 100% */
        color.opacity = (typeof swatch.opacity !== "undefined") ? swatch.opacity : 100;
        return color;
    }

    // =========================================
    // CMYK の自動生成 / CMYK color generation
    // =========================================

    /**
     * 範囲内の整数の乱数を返す
     * @param {number} min - 最小値
     * @param {number} max - 最大値
     * @returns {number} min 以上 max 以下の整数
     */
    function randInt(min, max) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    /**
     * 配列をシャッフルした複製を返す
     * @param {Array} sourceArray - 元の配列
     * @returns {Array} シャッフルした配列
     */
    function shuffleArray(sourceArray) {
        var result = sourceArray.slice();
        for (var i = result.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = result[i];
            result[i] = result[j];
            result[j] = temp;
        }
        return result;
    }

    /**
     * 既存の色すべてから、CMY のマンハッタン距離で十分に離れているか
     * @param {number[]} cmy - [C, M, Y]
     * @param {CMYKColor[]} existingColors - 採用済みの色
     * @param {number} minDistance - 最小距離
     * @returns {boolean} 離れていれば true
     */
    function isFarEnoughCMY(cmy, existingColors, minDistance) {
        for (var i = 0; i < existingColors.length; i++) {
            var existing = existingColors[i];
            var distance = Math.abs(cmy[0] - existing.cyan) + Math.abs(cmy[1] - existing.magenta) + Math.abs(cmy[2] - existing.yellow);
            if (distance < minDistance) {
                return false;
            }
        }
        return true;
    }

    /**
     * 2チャンネルだけ（K=0）の CMY 値をランダムに作る
     * @param {string} inkPair - "CM" / "CY" / "MY"
     * @param {number} maxTotal - 2チャンネルの合計の上限
     * @returns {number[]|null} [C, M, Y]（作れなければ null）
     */
    function randomTwoInkCMY(inkPair, maxTotal) {
        var first = randInt(1, Math.min(100, maxTotal - 1));
        var secondMax = Math.min(100, maxTotal - first);
        if (secondMax < 1) return null;
        var second = randInt(1, secondMax);

        if (inkPair === "CM") return [first, second, 0];
        if (inkPair === "CY") return [first, 0, second];
        return [0, first, second];
    }

    /**
     * CMYK ドキュメント用に、CM／CY／MY の2色混合（K=0）をできるだけ重複・近似なしで生成する
     * 数が非常に多く組み合わせが尽きたときは重複を許す
     * @param {number} count - 必要な色の数
     * @param {number} maxTotal - 2チャンネルの合計の上限
     * @returns {CMYKColor[]} 生成した色
     */
    function generateRandomCMYPaletteUnique(count, maxTotal) {
        var result = [];
        var seenKeys = {};
        var inkPairs = ["CM", "CY", "MY"];
        var minDistance = CMYK_FALLBACK_MIN_DISTANCE;

        /**
         * 並びが固定されないよう、ときどきペアの順をシャッフルして次のペアを選ぶ
         * @param {number} counter - 試行回数
         * @returns {string} 次に使うペア
         */
        function pickInkPair(counter) {
            if ((counter % 37) === 0) {
                inkPairs = shuffleArray(inkPairs);
            }
            return inkPairs[result.length % inkPairs.length];
        }

        /* 重複も近似も避けて生成（詰まったら距離制約を少しずつ緩める） / Avoid duplicates and near colors; relax the distance when stuck */
        var maxTries = Math.max(1500, count * 80);
        var tries = 0;
        var cmy, key;
        while (result.length < count && tries++ < maxTries) {
            if (minDistance > 0 && (tries % 500) === 0) {
                minDistance = Math.max(0, minDistance - 5);
            }
            cmy = randomTwoInkCMY(pickInkPair(tries), maxTotal);
            if (!cmy) continue;
            key = cmy.join(",");
            if (seenKeys[key]) continue;
            if (minDistance > 0 && !isFarEnoughCMY(cmy, result, minDistance)) continue;
            seenKeys[key] = true;
            result.push(createCMYKColor([cmy[0], cmy[1], cmy[2], 0]));
        }

        /* 埋まらないときは重複を許す（距離はベストエフォート、3000回で打ち切り） / Allow duplicates; distance is best effort */
        var guard = 0;
        while (result.length < count) {
            if (guard++ > 3000) {
                minDistance = 0;
            }
            cmy = randomTwoInkCMY(pickInkPair(guard), maxTotal);
            if (!cmy) continue;
            if (minDistance > 0 && !isFarEnoughCMY(cmy, result, minDistance)) continue;
            result.push(createCMYKColor([cmy[0], cmy[1], cmy[2], 0]));
        }

        return result;
    }

    // =========================================
    // 色付け / Coloring
    // =========================================

    /**
     * 使う色の一覧を決める（スウォッチの選択が1色以下・白のみなら定義済みカラーや自動生成）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} colorTargets - 色付け対象
     * @param {TextRange|null} selectedTextRange - 選択中の文字範囲
     * @returns {Array<{color: Color}>} スウォッチ（または color を持つオブジェクト）
     */
    function resolveSwatches(doc, colorTargets, selectedTextRange) {
        var selectedSwatches = doc.swatches.getSelected();
        if (selectedSwatches && selectedSwatches.length > 1 && !allWhiteSwatches(selectedSwatches)) {
            return selectedSwatches;
        }

        var isCMYK = (doc.documentColorSpace === DocumentColorSpace.CMYK);
        var fallbackColors;
        if (colorTargets.length === 4 && !selectedTextRange) {
            fallbackColors = getFourColorPreset(doc.documentColorSpace);
        } else if (isCMYK) {
            /* 対象の数だけ生成 / Generate one color per target */
            fallbackColors = generateRandomCMYPaletteUnique(getNeededColorCount(colorTargets, selectedTextRange), CMYK_FALLBACK_MAX_TOTAL);
        } else {
            fallbackColors = createColors(RGB_FALLBACK_COLORS, createRGBColor);
        }

        var fallbackSwatches = [];
        for (var i = 0; i < fallbackColors.length; i++) {
            fallbackSwatches.push({ color: fallbackColors[i] });
        }
        return fallbackSwatches;
    }

    /**
     * 文字ごとに色を付け、線をなし・不透明度を100%にする
     * @param {Characters} characters - 文字のコレクション
     * @param {number} charCount - 色を付ける文字数
     * @param {Array<{color: Color}>} swatches - 使う色
     * @returns {void}
     */
    function colorCharacters(characters, charCount, swatches) {
        for (var i = 0; i < charCount; i++) {
            characters[i].fillColor = getSwatchColor(i, swatches);
            characters[i].strokeColor = new NoColor();
            characters[i].opacity = 100;
        }
    }

    /**
     * パスに色を付け、線をなし・不透明度を100%にする
     * @param {PathItem} pathItem - 対象のパス
     * @param {Color} color - 塗りの色
     * @returns {void}
     */
    function colorPath(pathItem, color) {
        pathItem.fillColor = color;
        pathItem.stroked = false;
        pathItem.opacity = 100;
    }

    /**
     * オブジェクトを位置順に並べ替える（横に広ければ左→右、縦に広ければ上→下）
     * @param {PageItem[]} items - 並べ替えるオブジェクト（その場で並べ替える）
     * @returns {void}
     */
    function sortByPosition(items) {
        var hMin = Infinity, hMax = -Infinity, vMin = Infinity, vMax = -Infinity;
        for (var i = 0; i < items.length; i++) {
            var left = items[i].left;
            var top = items[i].top;
            if (left < hMin) hMin = left;
            if (left > hMax) hMax = left;
            if (top < vMin) vMin = top;
            if (top > vMax) vMax = top;
        }
        if (hMax - hMin > vMax - vMin) {
            items.sort(function (a, b) { return comparePosition(a.left, b.left, b.top, a.top); });
        } else {
            items.sort(function (a, b) { return comparePosition(b.top, a.top, a.left, b.left); });
        }
    }

    /**
     * 主キーで比べ、同じなら副キーで比べる
     * @param {number} primaryA - 主キー（a）
     * @param {number} primaryB - 主キー（b）
     * @param {number} secondaryA - 副キー（a）
     * @param {number} secondaryB - 副キー（b）
     * @returns {number} 比較結果
     */
    function comparePosition(primaryA, primaryB, secondaryA, secondaryB) {
        return primaryA == primaryB ? secondaryA - secondaryB : primaryA - primaryB;
    }

    /**
     * 複数のオブジェクトを位置順に並べ、1つずつ色を付ける
     * @param {PageItem[]} colorTargets - 色付け対象
     * @param {Array<{color: Color}>} swatches - 使う色
     * @returns {void}
     */
    function colorItemsByPosition(colorTargets, swatches) {
        sortByPosition(colorTargets);
        for (var i = 0; i < colorTargets.length; i++) {
            var swatchColor = getSwatchColor(i, swatches);
            var currentItem = colorTargets[i];

            if (currentItem.typename === "PathItem") {
                colorPath(currentItem, swatchColor);
            } else if (currentItem.typename === "CompoundPathItem" && currentItem.pathItems.length > 0) {
                for (var j = 0; j < currentItem.pathItems.length; j++) {
                    colorPath(currentItem.pathItems[j], swatchColor);
                }
            } else if (currentItem.typename === "TextFrame") {
                /* 複数テキストはテキスト単位で色付け / Color each text frame as a whole */
                currentItem.textRange.fillColor = swatchColor;
                currentItem.textRange.strokeColor = new NoColor();
                currentItem.textRange.opacity = 100;
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のオブジェクトやテキストに色を付ける
     * @returns {void}
     */
    function main() {
        /* 予期しない DOM エラーはまとめて通知 / Report unexpected DOM errors */
        try {
            if (app.documents.length === 0) {
                alert(getLabel("alert.noDocument"));
                return;
            }
            var doc = app.activeDocument;
            /* グループ内のテキスト・パスも対象にする / Include text and paths inside groups */
            var colorTargets = flattenSelection(doc.selection);
            /* テキスト編集中の文字範囲 / Character range while editing text */
            var selectedTextRange = getSingleSelectedTextRange(doc.selection);

            if (colorTargets.length === 0 && !selectedTextRange) {
                alert(getLabel("alert.noSelection"));
                return;
            }

            var swatches = resolveSwatches(doc, colorTargets, selectedTextRange);
            /* 3色より多ければシャッフル / Shuffle when there are more than three colors */
            if (swatches.length > 3) {
                swatches = shuffleArray(swatches);
            }

            if (selectedTextRange) {
                colorCharacters(selectedTextRange.characters, selectedTextRange.characters.length, swatches);
            } else if (isSingleTextFrame(colorTargets)) {
                colorCharacters(colorTargets[0].characters, colorTargets[0].contents.length, swatches);
            } else {
                colorItemsByPosition(colorTargets, swatches);
            }
        } catch (e) {
            alert(getLabel("alert.unexpected") + e.message);
        }
    }

    main();

})();
