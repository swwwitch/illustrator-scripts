#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの不透明度を、塗りのカラーそのものに焼き込んで不透明にします。
親グループの不透明度も再帰的に合成し、重なったオブジェクトは背面から合成して見た目の色を再現します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenOpacityPro.md

### Overview

Bakes the opacity of the selected objects into their fill colors so that everything becomes fully opaque.
Parent group opacity is composited recursively, and overlapping objects are composited from the back to reproduce the apparent color.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenOpacityPro.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenOpacityPro";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenOpacityPro.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenOpacityPro.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「同じ形」とみなす許容値。パスファインダー・分割後の小さな誤差を丸める
       Tolerances for grouping "same shape"; quantize to absorb float jitter after Pathfinder/Expand */
    var GEOM_TOL_PT = 0.1;        /* 位置・サイズの許容値（pt） / position/size tolerance in points */
    var AREA_TOL = 0.01;          /* 面積の許容値（PathItem.area は平方ポイント） / area tolerance (PathItem.area is in square points) */

    /* true: RGB に変換してリニア光で合成する（画面上の不透明度の見え方に近いが、カラープロファイルの変換で色がずれることがある）
       false: 元のカラースペースのまま合成する（従来の挙動。元の色相を保ちやすい）
       If true: convert to RGB and blend in linear light (closer to on-screen compositing, but profile conversion may shift colors).
       If false: blend directly in the source color space (previous behavior; preserves original hues better). */
    var USE_GAMMA_CORRECT_BLEND = false;

    /* true: 単体を白と合成するとき RGB（リニア光）で合成する。CMYK のインキを単純に掛けたときの「暗すぎる」結果を避ける
       If true: when flattening a single item (opacity over white), composite in RGB (linear light)
       to avoid the common "too dark" result from naive CMYK ink scaling. */
    var USE_RGB_WHITE_COMPOSITE = false;

    // =========================================
    // 形の判定 / Shape keys
    // =========================================

    /**
     * 値を刻み幅の整数（tick）に丸める。浮動小数の文字列化の揺れを避けるため、形のキーは tick で作る
     * @param {number} value - 元の値
     * @param {number} step - 刻み幅
     * @returns {number} tick（step が 0 以下なら value のまま）
     */
    function toTick(value, step) {
        if (!step || step <= 0) return value;
        return Math.round(value / step);
    }

    /**
     * 親（グループまたはレイヤー）の中での重なり順の番号を返す（0 が最前面）
     * zOrderPosition よりバージョン間で確実
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {number|null} 番号（求められないときは null）
     */
    function getStackIndexInParent(targetItem) {
        if (!targetItem) return null;
        try {
            var parentContainer = targetItem.parent;
            if (!parentContainer) return null;
            var siblingItems = parentContainer.pageItems;
            if (!siblingItems) return null;
            for (var i = 0; i < siblingItems.length; i++) {
                if (siblingItems[i] === targetItem) return i;
            }
        } catch (e) { /* 親や兄弟を読めないときは不明とする / unknown when the parent cannot be read */ }
        return null;
    }

    /**
     * コンテナ内のパスを再帰的に集める（グループの中もたどる）
     * @param {Object} container - GroupItem など pageItems を持つもの
     * @param {PathItem[]} outPathItems - パスを追加する配列
     * @returns {void}
     */
    function getAllPathItems(container, outPathItems) {
        var childItems = container.pageItems;
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            if (childItem.typename === "PathItem") {
                outPathItems.push(childItem);
            } else if (childItem.typename === "GroupItem") {
                getAllPathItems(childItem, outPathItems);
            }
        }
    }

    /**
     * 分割後のパスから、形のキー・不透明度・塗り・重なり順を読み出す（読めないパスは飛ばす）
     * @param {PathItem[]} dividedPaths - 分割後のパス（[0] が最前面）
     * @returns {Object[]} パスごとの情報
     */
    function collectPathEntries(dividedPaths) {
        var pathEntries = [];
        for (var i = 0; i < dividedPaths.length; i++) {
            var pathItem = dividedPaths[i];
            try {
                var bounds = pathItem.geometricBounds; /* [left, top, right, bottom] */
                var zOrderPos = null;
                try { zOrderPos = pathItem.zOrderPosition; } catch (e) { zOrderPos = null; }
                var stackIndex = getStackIndexInParent(pathItem);

                pathEntries.push({
                    pathItem: pathItem,

                    /* 向き（時計回り・反時計回り）で符号が変わらないよう絶対値の面積を tick に / Absolute area as ticks */
                    areaTick: toTick(Math.abs(pathItem.area), AREA_TOL),

                    /* left/top だけより安定する geometricBounds を tick に / geometricBounds as ticks */
                    leftTick: toTick(bounds[0], GEOM_TOL_PT),
                    topTick: toTick(bounds[1], GEOM_TOL_PT),
                    rightTick: toTick(bounds[2], GEOM_TOL_PT),
                    bottomTick: toTick(bounds[3], GEOM_TOL_PT),

                    pointCount: pathItem.pathPoints.length,
                    opacity: pathItem.opacity,
                    fillColor: pathItem.filled ? pathItem.fillColor : null,

                    /* 奥行き: 親の中の重なり順（0 が最前面）→ zOrderPosition → 取り出し順の順で使う。大きいほど背面
                       Depth: parent stacking index, then zOrderPosition, then extraction order. Larger means more back */
                    depth: (stackIndex !== null) ? stackIndex : ((zOrderPos !== null) ? zOrderPos : i)
                });
            } catch (e) {
                /* エラーが出るアイテムはスキップ / skip items that throw */
            }
        }
        return pathEntries;
    }

    /**
     * 同じ形（面積・境界・アンカー数が同じ）のパスごとにまとめる
     * @param {Object[]} pathEntries - collectPathEntries() の結果
     * @returns {Object} 形のキー → そのキーのパス情報の配列
     */
    function groupByGeometry(pathEntries) {
        var geometryGroups = {};
        for (var i = 0; i < pathEntries.length; i++) {
            var pathEntry = pathEntries[i];
            var geometryKey = pathEntry.areaTick + "_" + pathEntry.leftTick + "_" + pathEntry.topTick + "_" +
                pathEntry.rightTick + "_" + pathEntry.bottomTick + "_" + pathEntry.pointCount;
            if (!geometryGroups[geometryKey]) geometryGroups[geometryKey] = [];
            geometryGroups[geometryKey].push(pathEntry);
        }
        return geometryGroups;
    }

    // =========================================
    // 色の変換 / Color conversion
    // =========================================

    /* Illustrator の不透明度の合成は実質 RGB の表示空間で行われる。チャンネル値を直接合成すると
       （特に CMYK やガンマ補正なしの sRGB では）見た目より暗くなりがち。そのため RGB に変換して
       リニア光で合成し、ドキュメントのカラースペースに戻す方法も用意している
       Illustrator opacity compositing is effectively done in RGB display space. Blending channel values
       directly tends to look darker, so we can convert to RGB, blend in linear light, then convert back. */

    /**
     * 値を範囲内に収める
     * @param {number} value - 元の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, value));
    }

    /**
     * K だけの CMYK カラーを作る
     * @param {number} blackValue - K の値
     * @returns {CMYKColor} C・M・Y が 0 のカラー
     */
    function createKOnlyCMYK(blackValue) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = blackValue;
        return cmykColor;
    }

    /**
     * RGB 値（0〜255）から RGBColor を作る
     * @param {number[]} rgb8 - [r, g, b]
     * @returns {RGBColor} RGB カラー
     */
    function createRGBColor(rgb8) {
        var rgbColor = new RGBColor();
        rgbColor.red = rgb8[0];
        rgbColor.green = rgb8[1];
        rgbColor.blue = rgb8[2];
        return rgbColor;
    }

    /**
     * K だけの CMYK（無彩色）かを判定する
     * @param {Color} color - 判定する色
     * @returns {boolean} C・M・Y がほぼ 0 の CMYK なら true
     */
    function isNeutralCMYK(color) {
        if (!color || color.typename !== 'CMYKColor') return false;
        var tolerance = 1e-6;
        return (Math.abs(color.cyan) < tolerance && Math.abs(color.magenta) < tolerance && Math.abs(color.yellow) < tolerance);
    }

    /**
     * K だけの CMYK を sRGB のグレー値に写す（プロファイル変換なし。K=0 → 255 白、K=100 → 0 黒）
     * @param {CMYKColor} color - K だけの CMYK
     * @returns {number} グレー値（0〜255）
     */
    function neutralCMYKToGray255(color) {
        var blackValue = clampValue(color.black, 0, 100);
        return 255 * (1 - blackValue / 100);
    }

    /**
     * sRGB のグレー値を K だけの CMYK に戻す
     * @param {number} gray255 - グレー値（0〜255）
     * @returns {CMYKColor} K だけの CMYK
     */
    function gray255ToNeutralCMYK(gray255) {
        var grayValue = clampValue(gray255, 0, 255);
        var blackValue = 100 * (1 - grayValue / 255);
        return createKOnlyCMYK(clampValue(blackValue, 0, 100));
    }

    /**
     * sRGB の 8bit 値をリニア光（0〜1）にする（ガンマ 2.2 の近似）
     * @param {number} srgb8 - 0〜255
     * @returns {number} 0〜1
     */
    function srgb8ToLinear01(srgb8) {
        return Math.pow(srgb8 / 255, 2.2);
    }

    /**
     * リニア光（0〜1）を sRGB の 8bit 値にする
     * @param {number} linear01 - 0〜1
     * @returns {number} 0〜255 の整数
     */
    function linear01ToSrgb8(linear01) {
        return Math.round(Math.pow(clampValue(linear01, 0, 1), 1 / 2.2) * 255);
    }

    /**
     * 色を RGB（0〜255）の配列にする
     * @param {Color} color - RGB または CMYK の色
     * @returns {number[]|null} [r, g, b]（スポット・パターン・グラデーションなどは null）
     */
    function colorToRGB8(color) {
        if (!color) return null;

        if (color.typename === "RGBColor") {
            return [color.red, color.green, color.blue];
        }
        if (color.typename === "CMYKColor") {
            try {
                /* convertSampleColor は変換先の値域で返す（RGB: 0〜255） / returns values in the target range */
                var convertedRGB = app.convertSampleColor(
                    ImageColorSpace.CMYK,
                    [color.cyan, color.magenta, color.yellow, color.black],
                    ImageColorSpace.RGB,
                    ColorConvertPurpose.defaultpurpose
                );
                return [convertedRGB[0], convertedRGB[1], convertedRGB[2]];
            } catch (e) {
                /* 変換できないときの簡易換算 / naive fallback conversion */
                var red = 255 * (1 - color.cyan / 100) * (1 - color.black / 100);
                var green = 255 * (1 - color.magenta / 100) * (1 - color.black / 100);
                var blue = 255 * (1 - color.yellow / 100) * (1 - color.black / 100);
                return [red, green, blue];
            }
        }
        return null;
    }

    /**
     * RGB（0〜255）をドキュメントのカラースペースの色にする（CMYK ドキュメントなら CMYKColor、RGB なら RGBColor）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} rgb8 - [r, g, b]
     * @returns {Color|null} 色（rgb8 が無ければ null）
     */
    function rgb8ToDocColor(doc, rgb8) {
        if (!rgb8) return null;

        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            try {
                var convertedCMYK = app.convertSampleColor(
                    ImageColorSpace.RGB,
                    [rgb8[0], rgb8[1], rgb8[2]],
                    ImageColorSpace.CMYK,
                    ColorConvertPurpose.defaultpurpose
                );
                var cmykColor = new CMYKColor();
                cmykColor.cyan = convertedCMYK[0];
                cmykColor.magenta = convertedCMYK[1];
                cmykColor.yellow = convertedCMYK[2];
                cmykColor.black = convertedCMYK[3];
                return cmykColor;
            } catch (e) {
                /* 変換できないときは RGB のまま渡す（Illustrator が内部で変換する） / let Illustrator convert RGB */
                return createRGBColor(rgb8);
            }
        }
        return createRGBColor(rgb8);
    }

    // =========================================
    // 色の合成 / Color blending
    // =========================================

    /**
     * RGB 同士をリニア光で合成する（out = top * alpha + bottom * (1 - alpha)）
     * @param {number[]} topRGB8 - 前面の [r, g, b]
     * @param {number[]} bottomRGB8 - 背面の [r, g, b]
     * @param {number} alpha - 前面の不透明度（0〜1）
     * @returns {number[]} 合成した [r, g, b]（0〜255 の整数）
     */
    function blendRGB8Linear(topRGB8, bottomRGB8, alpha) {
        var inverseAlpha = 1.0 - alpha;
        var blended = [];
        for (var i = 0; i < 3; i++) {
            var topLinear = srgb8ToLinear01(topRGB8[i]);
            var bottomLinear = srgb8ToLinear01(bottomRGB8[i]);
            blended.push(linear01ToSrgb8(topLinear * alpha + bottomLinear * inverseAlpha));
        }
        return blended;
    }

    /**
     * 色を白の上にリニア光で合成し、ドキュメントのカラースペースの色にする
     * @param {Color} color - 合成する色
     * @param {number} alpha - 不透明度（0〜1）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Color|null} 合成した色（RGB に変換できない種類は null）
     */
    function compositeOverWhiteLinear(color, alpha, doc) {
        var rgb8 = colorToRGB8(color);
        if (!rgb8) return null;

        var clampedRGB8 = [clampValue(rgb8[0], 0, 255), clampValue(rgb8[1], 0, 255), clampValue(rgb8[2], 0, 255)];
        return rgb8ToDocColor(doc, blendRGB8Linear(clampedRGB8, [255, 255, 255], alpha));
    }

    /**
     * 重なりの最背面の色を、自身の不透明度で白の上に合成する（RGB リニア光。K だけの CMYK は K を掛けるだけ）
     * @param {Color} color - 最背面の塗り
     * @param {number} alpha - 不透明度（0〜1）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Color} 合成した色（不透明・未対応の種類は元の色）
     */
    function blendOverWhiteForOverlap(color, alpha, doc) {
        if (!color) return color;
        if (!doc) return color;

        /* 不透明なら余計な変換をしない / keep as-is when fully opaque */
        if (alpha >= 0.999) return color;

        /* CMYK の無彩色（K だけ）は RGB を経由せずに K を掛ける（例: K100 の 50% → K50）
           Neutral CMYK: avoid the RGB roundtrip that can distort K */
        if (doc.documentColorSpace === DocumentColorSpace.CMYK && isNeutralCMYK(color)) {
            return createKOnlyCMYK(color.black * alpha);
        }

        return compositeOverWhiteLinear(color, alpha, doc) || color;
    }

    /**
     * 前面の色を背面の色に不透明度で合成する
     * 既定は同じカラースペースのまま合成（USE_GAMMA_CORRECT_BLEND が true なら RGB リニア光）
     * @param {Color} topColor - 前面の色
     * @param {Color} bottomColor - 背面の色
     * @param {number} alpha - 前面の不透明度（0〜1）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Color} 合成した色（未対応の種類は前面の色）
     */
    function blendColors(topColor, bottomColor, alpha, doc) {
        var inverseAlpha = 1.0 - alpha;

        if (!topColor) return bottomColor;
        if (!bottomColor) return topColor;

        /* CMYK の無彩色同士はグレーの RGB で合成して重ねすぎの暗さを避ける / Neutral CMYK: blend as gray in linear light */
        if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK && isNeutralCMYK(topColor) && isNeutralCMYK(bottomColor)) {
            var topGray = neutralCMYKToGray255(topColor);
            var bottomGray = neutralCMYKToGray255(bottomColor);
            var blendedGray = blendRGB8Linear([topGray, topGray, topGray], [bottomGray, bottomGray, bottomGray], alpha);
            return gray255ToNeutralCMYK(blendedGray[0]);
        }

        if (!USE_GAMMA_CORRECT_BLEND) {
            if (topColor.typename === "CMYKColor" && bottomColor.typename === "CMYKColor") {
                var cmykColor = new CMYKColor();
                cmykColor.cyan = topColor.cyan * alpha + bottomColor.cyan * inverseAlpha;
                cmykColor.magenta = topColor.magenta * alpha + bottomColor.magenta * inverseAlpha;
                cmykColor.yellow = topColor.yellow * alpha + bottomColor.yellow * inverseAlpha;
                cmykColor.black = topColor.black * alpha + bottomColor.black * inverseAlpha;
                return cmykColor;
            } else if (topColor.typename === "RGBColor" && bottomColor.typename === "RGBColor") {
                var rgbColor = new RGBColor();
                rgbColor.red = Math.round(topColor.red * alpha + bottomColor.red * inverseAlpha);
                rgbColor.green = Math.round(topColor.green * alpha + bottomColor.green * inverseAlpha);
                rgbColor.blue = Math.round(topColor.blue * alpha + bottomColor.blue * inverseAlpha);
                return rgbColor;
            }
            /* 種類が違う・未対応なら前面の色 / keep the top color for other types */
            return topColor;
        }

        /* RGB リニア光で合成してドキュメントのカラースペースに戻す / gamma-correct blend in RGB */
        if (!doc) return topColor;

        var topRGB8 = colorToRGB8(topColor);
        var bottomRGB8 = colorToRGB8(bottomColor);
        if (!topRGB8 || !bottomRGB8) return topColor;

        return rgb8ToDocColor(doc, blendRGB8Linear(topRGB8, bottomRGB8, alpha));
    }

    /**
     * 色を白の上に不透明度で合成する（重なりの無い単体用）
     * @param {Color} color - 合成する色
     * @param {number} alpha - 不透明度（0〜1）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Color} 合成した色（未対応の種類は元の色）
     */
    function blendWithWhite(color, alpha, doc) {
        var inverseAlpha = 1.0 - alpha;
        if (!color) return color;

        if (USE_RGB_WHITE_COMPOSITE && doc) {
            var composited = compositeOverWhiteLinear(color, alpha, doc);
            if (composited) return composited;
            /* RGB にできない種類は下の従来の方法へ / fall through for unsupported types */
        }

        /* 従来の方法（色相は保つが、CMYK では見た目より暗くなることがある） / previous direct method */
        if (color.typename === "CMYKColor") {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = color.cyan * alpha;
            cmykColor.magenta = color.magenta * alpha;
            cmykColor.yellow = color.yellow * alpha;
            cmykColor.black = color.black * alpha;
            return cmykColor;
        } else if (color.typename === "RGBColor") {
            var rgbColor = new RGBColor();
            rgbColor.red = Math.round(color.red * alpha + 255 * inverseAlpha);
            rgbColor.green = Math.round(color.green * alpha + 255 * inverseAlpha);
            rgbColor.blue = Math.round(color.blue * alpha + 255 * inverseAlpha);
            return rgbColor;
        }
        return color;
    }

    // =========================================
    // 不透明度の焼き込み / Baking opacity
    // =========================================

    /**
     * オブジェクトの不透明度を 0〜1 の比率で返す
     * 種類によっては opacity を持たず例外になるため、その場合は 1（不透明）として扱う。
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number} 不透明度の比率（0〜1）
     */
    function getItemOpacityRatio(item) {
        try {
            return item.opacity / 100;
        } catch (e) {
            return 1;
        }
    }

    /**
     * オブジェクトの不透明度を 100% に戻す（設定できない種類は何もしない）
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {void}
     */
    function resetItemOpacity(item) {
        try {
            item.opacity = 100;
        } catch (e) {}
    }

    /**
     * 1つだけ選択したとき（重なりなし）に、親の不透明度も掛け合わせて塗りに焼き込む（再帰）
     * 分割・分割拡張のメニュー操作で透明が潰れて不透明度が失われるのを避ける。線はそのまま、テキストや配置画像は対象外
     * @param {PageItem} item - 対象のオブジェクト
     * @param {number} parentAlpha - 親から受け継ぐ不透明度（0〜1）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function bakeOpacityIntoFillRecursive(item, parentAlpha, doc) {
        if (!item) return;
        if (parentAlpha === undefined || parentAlpha === null) parentAlpha = 1.0;

        var typeName = item.typename;

        if (typeName === 'GroupItem' || typeName === 'CompoundPathItem') {
            var containerAlpha = parentAlpha * getItemOpacityRatio(item);
            try {
                var childItems = (typeName === 'GroupItem') ? item.pageItems : item.pathItems;
                for (var i = 0; i < childItems.length; i++) {
                    bakeOpacityIntoFillRecursive(childItems[i], containerAlpha, doc);
                }
            } catch (e) { }
            resetItemOpacity(item);
            return;
        }

        if (typeName === 'PathItem') {
            var pathAlpha = parentAlpha * getItemOpacityRatio(item);
            /* 塗りだけ焼き込み、線はそのまま / bake the fill only */
            try {
                if (item.filled && item.fillColor) {
                    item.fillColor = blendWithWhite(item.fillColor, pathAlpha, doc);
                }
            } catch (e) { }
            resetItemOpacity(item);
        }
    }

    /**
     * 同じ形で重なったパスを背面→前面の順に合成し、最前面の1つに最終色を適用して残りを削除する
     * 最前面を残すことで、周囲のオブジェクトとの前後関係を変えない
     * @param {Object[]} sameShapeEntries - 背面→前面の順に並べたパス情報
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function mergeOverlappingEntries(sameShapeEntries, doc) {
        var backEntry = sameShapeEntries[0];                             /* 最背面 / backmost */
        var survivorEntry = sameShapeEntries[sameShapeEntries.length - 1]; /* 最前面 / frontmost */

        /* 最背面は自身の不透明度で白の上に合成してから始める / start from the backmost composited over white */
        var baseColor = backEntry.fillColor;
        if (baseColor) {
            baseColor = blendOverWhiteForOverlap(baseColor, backEntry.opacity / 100, doc);
        }

        /* 背面→前面の順に合成 / Composite back to front */
        for (var i = 1; i < sameShapeEntries.length; i++) {
            var frontEntry = sameShapeEntries[i];
            if (baseColor && frontEntry.fillColor) {
                baseColor = blendColors(frontEntry.fillColor, baseColor, frontEntry.opacity / 100, doc);
            }
        }

        /* 最前面に最終色を適用して不透明度 100 に / Apply the final color to the frontmost */
        try {
            survivorEntry.pathItem.fillColor = baseColor;
        } catch (e) {}
        try {
            survivorEntry.pathItem.opacity = 100;
        } catch (e) {}

        /* 残りは削除（最前面は残して重なり順を保つ） / Remove the others */
        for (var j = 0; j < sameShapeEntries.length; j++) {
            if (sameShapeEntries[j] === survivorEntry) continue;
            try { sameShapeEntries[j].pathItem.remove(); } catch (e) {}
        }
    }

    /**
     * 同じ形のパスの組を平坦化する（重なりありなら合成、単品なら白と合成）
     * @param {Object[]} sameShapeEntries - 同じ形のパス情報
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function flattenGeometryGroup(sameShapeEntries, doc) {
        /* 奥行きの大きい順（背面→前面）に並べる / Sort back to front */
        sameShapeEntries.sort(function (a, b) { return b.depth - a.depth; });

        if (sameShapeEntries.length > 1) {
            mergeOverlappingEntries(sameShapeEntries, doc);
            return;
        }

        /* 重なりなし（単品） / No overlap */
        var singleEntry = sameShapeEntries[0];
        if (singleEntry.fillColor) {
            singleEntry.pathItem.fillColor = blendWithWhite(singleEntry.fillColor, singleEntry.opacity / 100, doc);
        }
        singleEntry.pathItem.opacity = 100;
    }

    /**
     * 選択に線のあるオブジェクトが含まれるかを調べる（最上位のオブジェクトだけ）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 線幅が 0 より大きい線があれば true
     */
    function selectionHasStroke(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.stroked && selectedItem.strokeWidth > 0) {
                return true;
            }
        }
        return false;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の不透明度を塗りの色に焼き込んで不透明にする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        if (doc.selection.length < 1) {
            alert("オブジェクトを選択してください。");
            return;
        }

        /* 1. 線が含まれる場合は「パスのアウトライン」を実行 / Outline strokes first */
        if (selectionHasStroke(doc.selection)) {
            app.executeMenuCommand('OffsetPath v22');
            /* 再選択（念のため） / Reselect just in case */
            try {
                var outlinedSelection = doc.selection;
                doc.selection = null;
                doc.selection = outlinedSelection;
            } catch (e) {}
        }

        /* 1つだけなら重なりが無いので、分割せずに直接焼き込む（分割で透明が潰れ、K100 の 50% が K100 になるのを避ける）
           Single object: bake directly to avoid Divide/Expand collapsing transparency */
        if (doc.selection && doc.selection.length === 1) {
            bakeOpacityIntoFillRecursive(doc.selection[0], 1.0, doc);
            alert('処理が完了しました。');
            return;
        }

        /* 2. 分割 / Divide */
        app.executeMenuCommand('group');
        app.executeMenuCommand('Live Pathfinder Divide');
        app.executeMenuCommand('expandStyle');

        app.redraw(); /* 描画を強制更新 / force a redraw */
        var workGroup = doc.selection[0];
        if (!workGroup || workGroup.typename !== "GroupItem") return;

        /* 3. パスの情報を読み出し、同じ形ごとにまとめて合成（[0] が最前面、[last] が最背面）
           Read path info, group by shape, and composite ([0] is frontmost) */
        var dividedPaths = [];
        getAllPathItems(workGroup, dividedPaths);
        var geometryGroups = groupByGeometry(collectPathEntries(dividedPaths));
        for (var geometryKey in geometryGroups) {
            flattenGeometryGroup(geometryGroups[geometryKey], doc);
        }

        alert("処理が完了しました。");
    }

    main();

})();
