#target illustrator

/*

### 概要

開いている2つのドキュメントのうち、最前面のドキュメントのガイドを、もう一方のドキュメントの
同じ番号・同じサイズのアートボードへ複製します。

詳細は README を参照してください。

### Overview

Copies the guides of the frontmost document into the other open document, matching artboards
by index and size.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyGuidesToOtherDocument";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-07";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================
var MATCH_TOLERANCE = 0.001;   /* 同じ座標・同じサイズとみなす誤差（pt） */
var SKIP_DUPLICATES = true;    /* 複製先に同じ座標のガイドがある場合は作らない */

// =========================================
// ラベル定義 / Labels
// =========================================
var LABELS = {
    alert: {
        needTwoDocuments: { ja: "ドキュメントをちょうど2つ開いた状態で実行してください。", en: "Open exactly two documents before running this script." },
        noGuides: { ja: "複製できるガイドがありません。", en: "No guides to copy." },
        copied: { ja: "{0}本のガイドを複製しました。", en: "Copied {0} guide(s)." },
        copiedWithSkip: {
            ja: "{0}本のガイドを複製しました。\n番号またはサイズが一致しないアートボード：{1}",
            en: "Copied {0} guide(s).\nArtboards skipped for a mismatched index or size: {1}"
        }
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
     * 最前面のドキュメント以外の、もう一方のドキュメントを返す
     * @param {Document} sourceDoc - 基準にするドキュメント
     * @returns {Document} 複製先のドキュメント
     */
    function findOtherDocument(sourceDoc) {
        for (var i = 0; i < app.documents.length; i++) {
            if (app.documents[i] !== sourceDoc) {
                return app.documents[i];
            }
        }

        return null;
    }

    /**
     * 2つのアートボードが同じサイズかどうかを判定する
     * @param {number[]} rectA - アートボードの矩形 [左, 上, 右, 下]
     * @param {number[]} rectB - アートボードの矩形 [左, 上, 右, 下]
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {boolean} 同じサイズなら true
     */
    function isSameArtboardSize(rectA, rectB, tolerance) {
        var widthA = rectA[2] - rectA[0];
        var heightA = rectA[1] - rectA[3];
        var widthB = rectB[2] - rectB[0];
        var heightB = rectB[1] - rectB[3];

        return (
            Math.abs(widthA - widthB) <= tolerance &&
            Math.abs(heightA - heightB) <= tolerance
        );
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
     * 指定した名前のレイヤーを取得する（なければ作成する）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 見つかった、または作成したレイヤー
     */
    function findOrCreateLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }

        var newLayer = doc.layers.add();
        newLayer.name = layerName;

        return newLayer;
    }

    /**
     * 座標をアートボードの位置差分だけずらす
     * @param {number[]} point - 座標 [x, y]
     * @param {number} offsetX - X方向の差分
     * @param {number} offsetY - Y方向の差分
     * @returns {number[]} ずらした座標 [x, y]
     */
    function offsetPoint(point, offsetX, offsetY) {
        return [point[0] + offsetX, point[1] + offsetY];
    }

    /**
     * ガイドを複製先のレイヤーへ作成する（同じ座標のガイドが既にある場合は作らない）
     * @param {PathItem} source - 元のガイド
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {number} offsetX - X方向の差分
     * @param {number} offsetY - Y方向の差分
     * @param {PathItem[]} existingGuides - 複製先に既にあるガイド
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {PathItem} 作成したガイド（重複で作らなかった場合は null）
     */
    function copyGuideToLayer(source, targetLayer, offsetX, offsetY, existingGuides, tolerance) {
        var layerWasLocked = targetLayer.locked;

        if (layerWasLocked) {
            targetLayer.locked = false;
        }

        var newPath = targetLayer.pathItems.add();
        var sourcePoints = source.pathPoints;
        var pointCount = sourcePoints.length;

        for (var i = 0; i < pointCount; i++) {
            var sourcePoint = sourcePoints[i];
            var newPoint = newPath.pathPoints.add();
            newPoint.anchor = offsetPoint(sourcePoint.anchor, offsetX, offsetY);
            newPoint.leftDirection = offsetPoint(sourcePoint.leftDirection, offsetX, offsetY);
            newPoint.rightDirection = offsetPoint(sourcePoint.rightDirection, offsetX, offsetY);
            newPoint.pointType = sourcePoint.pointType;
        }

        newPath.closed = source.closed;

        var isDuplicated = false;

        if (SKIP_DUPLICATES) {
            for (var j = 0; j < existingGuides.length; j++) {
                if (isSameGeometry(newPath, existingGuides[j], tolerance)) {
                    isDuplicated = true;
                    break;
                }
            }
        }

        if (isDuplicated) {
            newPath.remove();
            newPath = null;
        } else {
            newPath.guides = true;
        }

        if (layerWasLocked) {
            targetLayer.locked = true;
        }

        return newPath;
    }

    if (app.documents.length !== 2) {
        alert(getLabel(LABELS.alert.needTwoDocuments));
        return;
    }

    var sourceDoc = app.activeDocument;
    var targetDoc = findOtherDocument(sourceDoc);
    var artboardCount = sourceDoc.artboards.length;
    var copiedCount = 0;
    var skippedArtboardCount = 0;

    for (var i = 0; i < artboardCount; i++) {
        /* 番号が対応するアートボードがなければ対象外 / Skip when the index has no counterpart */
        if (i >= targetDoc.artboards.length) {
            skippedArtboardCount++;
            continue;
        }

        var sourceRect = sourceDoc.artboards[i].artboardRect;
        var targetRect = targetDoc.artboards[i].artboardRect;

        /* サイズが違うアートボードは対象外 / Skip artboards with a different size */
        if (!isSameArtboardSize(sourceRect, targetRect, MATCH_TOLERANCE)) {
            skippedArtboardCount++;
            continue;
        }

        var sourceGuides = collectGuidesOnArtboard(sourceDoc, sourceRect);

        if (sourceGuides.length === 0) {
            continue;
        }

        /* アートボードの位置差分を求めて、同じ相対位置に複製 / Match the position within the artboard */
        var offsetX = targetRect[0] - sourceRect[0];
        var offsetY = targetRect[1] - sourceRect[1];
        var existingGuides = collectGuidesOnArtboard(targetDoc, targetRect);

        for (var j = 0; j < sourceGuides.length; j++) {
            var sourceGuide = sourceGuides[j];
            var targetLayer = findOrCreateLayer(targetDoc, sourceGuide.layer.name);
            var createdGuide = copyGuideToLayer(
                sourceGuide,
                targetLayer,
                offsetX,
                offsetY,
                existingGuides,
                MATCH_TOLERANCE
            );

            if (createdGuide !== null) {
                existingGuides.push(createdGuide);
                copiedCount++;
            }
        }
    }

    app.redraw();

    if (copiedCount === 0) {
        alert(getLabel(LABELS.alert.noGuides));
        return;
    }

    if (skippedArtboardCount > 0) {
        alert(
            getLabel(LABELS.alert.copiedWithSkip)
                .replace("{0}", copiedCount)
                .replace("{1}", skippedArtboardCount)
        );
        return;
    }

    alert(getLabel(LABELS.alert.copied).replace("{0}", copiedCount));
})();
