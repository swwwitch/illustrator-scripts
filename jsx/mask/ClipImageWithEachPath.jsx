#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ClipImageWithEachPath
Name=画像をパスごとに切り分け
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-11-03
Description-BEGIN
選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipImageWithEachPath";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-03";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "画像をパスごとに切り分け"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            needOneImage: { ja: "画像を1つとパスを選択してください（画像は1つだけ）。", en: "Select one image and the paths (only one image)." },
            noPaths: { ja: "マスクに使うパスを選択してください。", en: "Select the paths to use as masks." },
            noOverlap: { ja: "画像と重なっているパスがありません。", en: "None of the paths overlap the image." }
        }
    };

    // =========================================
    // 選択の仕分け / Sort the selection
    // =========================================

    /**
     * 画像（配置画像／埋め込み画像）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 画像なら true
     */
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    /**
     * マスクに使えるパス（パス／複合パス、ガイドは除く）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} マスクに使えれば true
     */
    function isMaskPath(item) {
        if (item.typename === "CompoundPathItem") return true;
        return item.typename === "PathItem" && !item.guides;
    }

    /**
     * 2つの境界（[左, 上, 右, 下]）が重なっているか
     * @param {number[]} boundsA - 1つめの境界
     * @param {number[]} boundsB - 2つめの境界
     * @returns {boolean} 重なっていれば true
     */
    function boundsOverlap(boundsA, boundsB) {
        return boundsA[0] < boundsB[2] && boundsB[0] < boundsA[2] &&
            boundsA[3] < boundsB[1] && boundsB[3] < boundsA[1];
    }

    // =========================================
    // パスごとにマスク / Clip with each path
    // =========================================

    /**
     * 画像の複製をパスでマスクし、パスのあった重ね順にクリップグループを置く
     * @param {PlacedItem|RasterItem} sourceImage - 元の画像
     * @param {PathItem|CompoundPathItem} maskPath - マスクに使うパス
     * @returns {GroupItem} 作ったクリップグループ
     */
    function clipImageWithPath(sourceImage, maskPath) {
        var clipGroup = maskPath.parent.groupItems.add();
        /* パスのすぐ前面にグループを置いてから中へ入れ、重ね順を保つ / Put the group right in front of the path, then move items in */
        clipGroup.move(maskPath, ElementPlacement.PLACEBEFORE);
        sourceImage.duplicate(clipGroup, ElementPlacement.PLACEATBEGINNING);
        maskPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;
        return clipGroup;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した画像を、選択したパスの数だけ複製してそれぞれマスクする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る / A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var selectedImages = [];
        var maskPaths = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (isImageItem(selectedItems[i])) selectedImages.push(selectedItems[i]);
            else if (isMaskPath(selectedItems[i])) maskPaths.push(selectedItems[i]);
        }
        if (selectedImages.length !== 1) {
            alert(getLabel(LABELS.alert.needOneImage));
            return;
        }
        if (maskPaths.length === 0) {
            alert(getLabel(LABELS.alert.noPaths));
            return;
        }

        /* 画像と重ならないパスは空のマスクになるので触らない / Leave paths outside the image alone; they would make empty masks */
        var sourceImage = selectedImages[0];
        var imageBounds = sourceImage.geometricBounds;
        var overlappingPaths = [];
        for (var j = 0; j < maskPaths.length; j++) {
            if (boundsOverlap(maskPaths[j].geometricBounds, imageBounds)) overlappingPaths.push(maskPaths[j]);
        }
        if (overlappingPaths.length === 0) {
            alert(getLabel(LABELS.alert.noOverlap));
            return;
        }

        doc.selection = null;
        var clipGroups = [];
        for (var k = 0; k < overlappingPaths.length; k++) {
            clipGroups.push(clipImageWithPath(sourceImage, overlappingPaths[k]));
        }
        sourceImage.remove();
        doc.selection = clipGroups;
    }

    main();

})();
