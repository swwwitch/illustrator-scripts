#target illustrator

/*

### 概要

アクティブアートボードのガイドを線付きのパスに変換してクリップボードへ送ります。
元のガイドは残し、同じ座標に重なった余分なガイドだけを削除します。

詳細は README を参照してください。

### Overview

Converts the guides on the active artboard into stroked paths and sends them to the clipboard.
The original guides are kept, while duplicated guides stacked at the same position are removed.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyGuidesAsPaths";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-07";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyGuidesAsPaths.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyGuidesAsPaths.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================
var STROKE_WIDTH    = 0.5;                       /* 変換後のパスに付ける線幅（pt） */
var MATCH_TOLERANCE = 0.001;                     /* 同じ座標とみなす誤差（pt） */
var TEMP_LAYER_NAME = "__guide_to_path__";       /* 作業用レイヤー名 */

// =========================================
// ラベル定義 / Labels
// =========================================
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noGuides: { ja: "対象となるガイドがありません。", en: "No guides found on the active artboard." },
        copied: { ja: "{0}本のガイドをコピーしました。", en: "Copied {0} guide(s)." }
    }
};

var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/**
 * ラベル定義から現在の言語の文言を取得する
 * @param {Object} entry - ja / en を持つラベル定義
 * @returns {string} 現在の言語の文言
 */
function getLabel(entry) {
    return entry[uiLang] || entry.en;
}

(function () {
    /**
     * ドキュメントのカラーモードに合わせた黒の線色を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} RGBColor または CMYKColor
     */
    function createStrokeColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        }

        var rgb = new RGBColor();
        rgb.red = 0;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * アートボード内にあるガイドを集める（ロックされたオブジェクト・レイヤーも対象）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {PathItem[]} 対象のガイド
     */
    function collectGuidesOnArtboard(doc, artboardRect) {
        var left = artboardRect[0];
        var top = artboardRect[1];
        var right = artboardRect[2];
        var bottom = artboardRect[3];
        var pathItems = doc.pathItems;
        var itemCount = pathItems.length;
        var result = [];

        for (var i = 0; i < itemCount; i++) {
            var item = pathItems[i];

            if (!item.guides) {
                continue;
            }

            /* ロックは無視し、非表示のものだけ除外 / Ignore lock, skip hidden only */
            if (item.hidden || !item.layer.visible) {
                continue;
            }

            var bounds = item.geometricBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            var centerY = (bounds[1] + bounds[3]) / 2;

            if (
                centerX >= left &&
                centerX <= right &&
                centerY <= top &&
                centerY >= bottom
            ) {
                result.push(item);
            }
        }

        return result;
    }

    /**
     * 2本のパスが同じ座標かどうかを判定する（描画方向の違いも同一とみなす）
     * @param {PathItem} pathA - 比較するパス
     * @param {PathItem} pathB - 比較するパス
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {boolean} 同じ座標なら true
     */
    function isSameGeometry(pathA, pathB, tolerance) {
        var pointsA = pathA.pathPoints;
        var pointsB = pathB.pathPoints;

        if (pathA.closed !== pathB.closed || pointsA.length !== pointsB.length) {
            return false;
        }

        var pointCount = pointsA.length;
        var forwardMatched = true;
        var reversedMatched = true;

        for (var i = 0; i < pointCount; i++) {
            var anchorA = pointsA[i].anchor;

            if (forwardMatched) {
                var anchorForward = pointsB[i].anchor;
                if (
                    Math.abs(anchorA[0] - anchorForward[0]) > tolerance ||
                    Math.abs(anchorA[1] - anchorForward[1]) > tolerance
                ) {
                    forwardMatched = false;
                }
            }

            if (reversedMatched) {
                var anchorReversed = pointsB[pointCount - 1 - i].anchor;
                if (
                    Math.abs(anchorA[0] - anchorReversed[0]) > tolerance ||
                    Math.abs(anchorA[1] - anchorReversed[1]) > tolerance
                ) {
                    reversedMatched = false;
                }
            }

            if (!forwardMatched && !reversedMatched) {
                return false;
            }
        }

        return true;
    }

    /**
     * ロックを一時的に解除してページアイテムを削除する
     * @param {PageItem} item - 削除するアイテム
     * @returns {void}
     */
    function removeLockedItem(item) {
        var parentLayer = item.layer;
        var layerWasLocked = parentLayer.locked;

        if (layerWasLocked) {
            parentLayer.locked = false;
        }

        if (item.locked) {
            item.locked = false;
        }

        item.remove();

        if (layerWasLocked) {
            parentLayer.locked = true;
        }
    }

    /**
     * 同じ座標に重なったガイドを1本だけ残し、残りをドキュメントから削除する
     * @param {PathItem[]} items - 対象のガイド
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {PathItem[]} 重複を取り除いたガイド
     */
    function removeDuplicatedGuides(items, tolerance) {
        var kept = [];

        for (var i = 0; i < items.length; i++) {
            var isDuplicated = false;

            for (var j = 0; j < kept.length; j++) {
                if (isSameGeometry(items[i], kept[j], tolerance)) {
                    isDuplicated = true;
                    break;
                }
            }

            if (isDuplicated) {
                removeLockedItem(items[i]);
            } else {
                kept.push(items[i]);
            }
        }

        return kept;
    }

    /**
     * ガイドの形状を写し取り、線付きの通常パスとして作業用レイヤーに作る
     * @param {PathItem} source - 元のガイド
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {Object} strokeColor - 適用する線色
     * @param {number} strokeWidth - 適用する線幅（pt）
     * @returns {PathItem} 作成した通常パス
     */
    function createStrokedCopy(source, targetLayer, strokeColor, strokeWidth) {
        var newPath = targetLayer.pathItems.add();
        var sourcePoints = source.pathPoints;
        var pointCount = sourcePoints.length;

        for (var i = 0; i < pointCount; i++) {
            var sourcePoint = sourcePoints[i];
            var newPoint = newPath.pathPoints.add();
            newPoint.anchor = sourcePoint.anchor;
            newPoint.leftDirection = sourcePoint.leftDirection;
            newPoint.rightDirection = sourcePoint.rightDirection;
            newPoint.pointType = sourcePoint.pointType;
        }

        newPath.closed = source.closed;
        newPath.filled = false;
        newPath.stroked = true;
        newPath.strokeWidth = strokeWidth;
        newPath.strokeColor = strokeColor;

        return newPath;
    }

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var targetGuides = collectGuidesOnArtboard(doc, artboard.artboardRect);

    if (targetGuides.length === 0) {
        alert(getLabel(LABELS.alert.noGuides));
        return;
    }

    /* 同じ座標に重なったガイドを削除 / Remove guides stacked at the same position */
    targetGuides = removeDuplicatedGuides(targetGuides, MATCH_TOLERANCE);

    /* 作業用レイヤーを作り、そこに線付きパスを作成 / Build stroked paths on a temp layer */
    var previousActiveLayer = doc.activeLayer;
    var tempLayer = doc.layers.add();
    tempLayer.name = TEMP_LAYER_NAME;

    var strokeColor = createStrokeColor(doc);
    var createdPaths = [];

    for (var i = 0; i < targetGuides.length; i++) {
        createdPaths.push(
            createStrokedCopy(targetGuides[i], tempLayer, strokeColor, STROKE_WIDTH)
        );
    }

    /* 作成したパスだけを選択 / Select only the new paths */
    app.executeMenuCommand("deselectall");

    for (var j = 0; j < createdPaths.length; j++) {
        createdPaths[j].selected = true;
    }

    /* クリップボードへカット / Cut to clipboard */
    app.redraw();                   /* 作成直後は再描画しないとカット対象にならない */
    app.executeMenuCommand("cut");  /* app.cut() は黙って無視されることがある */
    app.redraw();

    /* 作業用レイヤーを片付け、元の状態に戻す / Clean up the temp layer */
    tempLayer.remove();
    doc.activeLayer = previousActiveLayer;

    alert(getLabel(LABELS.alert.copied).replace("{0}", targetGuides.length));
})();
