#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数の画像と図形を選択して実行すると、対応する図形の大きさに合わせて各画像を拡大・縮小し、それぞれクリッピングマスクを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImgFitMaskMultiple.md

### Overview

With several images and shapes selected, scales each image to its matching shape and creates a clipping mask for each pair.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImgFitMaskMultiple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImgFitMaskMultiple";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImgFitMaskMultiple.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImgFitMaskMultiple.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectMultiple: { ja: "エラー: 複数の画像と図形を選択してください。", en: "Error: Select several images and shapes." },
            needShapeAndImage: {
                ja: "画像と図形がそれぞれ少なくとも1つずつ必要です。",
                en: "At least one image and one shape are required."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.noDocument" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    // =========================================
    // 画像とマスク / Image and mask
    // =========================================

    /**
     * マスクに使える図形（パス／複合パス）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 図形なら true
     */
    function isMaskShape(item) {
        return item.typename === "PathItem" || item.typename === "CompoundPathItem";
    }

    /**
     * 画像（配置画像／埋め込み画像）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 画像なら true
     */
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    /**
     * アイテムの中心座標を返す
     * @param {PageItem} item - 対象アイテム
     * @returns {number[]} [x, y]
     */
    function getCenter(item) {
        return [
            item.left + item.width / 2,
            item.top - item.height / 2
        ];
    }

    /**
     * 画像を図形を隙間なく覆う大きさに拡大・縮小して中央に合わせ、図形でクリッピングマスクを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem|RasterItem} image - マスクされる画像
     * @param {PathItem|CompoundPathItem} maskShape - マスクになる図形
     * @returns {GroupItem} 作成したクリップグループ
     */
    function fitImageToMask(doc, image, maskShape) {
        /* 隙間が出ないよう、倍率の大きい方を使う（全体を収めるなら Math.min）/ Use the larger ratio so no gap remains */
        var scaleFactor = Math.max(maskShape.width / image.width, maskShape.height / image.height);
        /* resize はパーセント指定 / resize() takes percentages */
        image.resize(scaleFactor * 100, scaleFactor * 100, true, true, true, true);

        /* リサイズ後の中心を図形の中心に合わせる / Center the resized image on the shape */
        var maskCenter = getCenter(maskShape);
        var imageCenter = getCenter(image);
        image.translate(maskCenter[0] - imageCenter[0], maskCenter[1] - imageCenter[1]);

        /* 図形の重ね順の位置にグループを作り、図形を最前面に入れる / Group at the shape's stacking position, shape on top */
        var clipGroup = doc.groupItems.add();
        clipGroup.move(maskShape, ElementPlacement.PLACEBEFORE);
        maskShape.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        image.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;
        return clipGroup;
    }

    // =========================================
    // 組み合わせ / Pairing
    // =========================================

    /**
     * 2点間の距離を返す
     * @param {number[]} pointA - [x, y]
     * @param {number[]} pointB - [x, y]
     * @returns {number} 距離
     */
    function getDistance(pointA, pointB) {
        return Math.sqrt(Math.pow(pointB[0] - pointA[0], 2) + Math.pow(pointB[1] - pointA[1], 2));
    }

    /**
     * 各画像に中心が最も近い図形を組み合わせ、距離の近い順に並べる
     * @param {PageItem[]} images - 画像
     * @param {PageItem[]} maskShapes - 図形
     * @returns {Object[]} { image, mask, distance } の配列（距離の昇順）
     */
    function buildNearestPairs(images, maskShapes) {
        var pairs = [];
        for (var i = 0; i < images.length; i++) {
            var imageCenter = getCenter(images[i]);
            var bestMask = null;
            var minDistance = Infinity;
            for (var j = 0; j < maskShapes.length; j++) {
                var distance = getDistance(imageCenter, getCenter(maskShapes[j]));
                if (distance < minDistance) {
                    minDistance = distance;
                    bestMask = maskShapes[j];
                }
            }
            if (bestMask) {
                pairs.push({ image: images[i], mask: bestMask, distance: minDistance });
            }
        }

        /* 近い組から処理し、遠くの誤判定を防ぐ / Nearest pairs first, to avoid distant mismatches */
        pairs.sort(function (a, b) {
            return a.distance - b.distance;
        });
        return pairs;
    }

    /**
     * 配列にアイテムが含まれているか
     * @param {PageItem[]} items - 配列
     * @param {PageItem} targetItem - 探すアイテム
     * @returns {boolean} 含まれていれば true
     */
    function containsItem(items, targetItem) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === targetItem) return true;
        }
        return false;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した画像と図形を近いもの同士で組み合わせ、それぞれクリッピングマスクを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;

        if (currentSelection.length < 2) {
            alert(getLabel("alert.selectMultiple"));
            return;
        }

        /* 選択を画像と図形に分ける / Split the selection into images and shapes */
        var images = [];
        var maskShapes = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var item = currentSelection[i];
            if (isMaskShape(item)) {
                maskShapes.push(item);
            } else if (isImageItem(item)) {
                images.push(item);
            }
        }

        if (images.length === 0 || maskShapes.length === 0) {
            alert(getLabel("alert.needShapeAndImage"));
            return;
        }

        /* 同じ図形を二度使わない / Never use the same shape twice */
        var pairs = buildNearestPairs(images, maskShapes);
        var usedMasks = [];
        for (var k = 0; k < pairs.length; k++) {
            if (containsItem(usedMasks, pairs[k].mask)) continue;
            fitImageToMask(doc, pairs[k].image, pairs[k].mask);
            usedMasks.push(pairs[k].mask);
        }
    }

    main();

})();
