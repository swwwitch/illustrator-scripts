#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

1つの図形と1つの画像を選択して実行すると、図形の大きさに合わせて画像を拡大・縮小してからクリッピングマスクを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImgFitMask.md

### Overview

With one shape and one image selected, scales the image to the shape and then creates a clipping mask.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImgFitMask.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImgFitMask";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImgFitMask.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImgFitMask.md"; /* README (English) */

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
            selectTwo: {
                ja: "エラー: 1つの図形（マスク用）と1つの画像を選択してください。",
                en: "Error: Select one shape (for the mask) and one image."
            },
            needShapeAndImage: {
                ja: "エラー: 「パス（図形）」と「配置画像」をそれぞれ1つずつ選択してください。",
                en: "Error: Select one path (shape) and one placed image."
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
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した図形と画像から、画像を図形に合わせたクリッピングマスクを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;

        if (currentSelection.length !== 2) {
            alert(getLabel("alert.selectTwo"));
            return;
        }

        var maskShape = null;   /* マスクになる図形 / shape used as the mask */
        var targetImage = null; /* マスクされる画像 / image being masked */
        for (var i = 0; i < currentSelection.length; i++) {
            var item = currentSelection[i];
            if (isMaskShape(item)) {
                maskShape = item;
            } else if (isImageItem(item)) {
                targetImage = item;
            }
        }

        if (!maskShape || !targetImage) {
            alert(getLabel("alert.needShapeAndImage"));
            return;
        }

        var clipGroup = fitImageToMask(doc, targetImage, maskShape);

        /* 作ったクリップグループを選択 / Select the new clipping group */
        doc.selection = null;
        clipGroup.selected = true;
    }

    main();

})();
