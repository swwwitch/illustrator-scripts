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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
