#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ClipImageWithEachPath
Name=画像をパスごとに切り分け
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-11-03
Description-BEGIN
選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。
選択に線（オープンパス）が含まれていると、画像の範囲をパスで区切り、区切られた領域ごとにマスクを作ります。

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.
When the selection includes lines (open paths), the image area is divided by the paths and each resulting region becomes a mask.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。
選択に線（オープンパス）が含まれていると、画像の範囲をパスで区切り、区切られた領域ごとにマスクを作ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.
When the selection includes lines (open paths), the image area is divided by the paths and each resulting region becomes a mask.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipImageWithEachPath";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-04";                   /* 更新日 / last updated */

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
            noOverlap: { ja: "画像と重なっているパスがありません。", en: "None of the paths overlap the image." },
            noRegions: { ja: "パスで区切られた領域を作れませんでした。", en: "Could not divide the image area with the paths." }
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
     * 区切り線として使う線（閉じていないパス）か
     * @param {PathItem|CompoundPathItem} item - 判定するパス
     * @returns {boolean} オープンパスなら true
     */
    function isDividerLine(item) {
        return item.typename === "PathItem" && !item.closed;
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
     * 画像の複製をパスでマスクし、指定したアイテムのすぐ前面にクリップグループを置く
     * @param {PlacedItem|RasterItem} sourceImage - 元の画像
     * @param {PathItem|CompoundPathItem} maskPath - マスクに使うパス
     * @param {PageItem} [placeAnchor] - この前面に置く（省略時はマスクに使うパス）
     * @returns {GroupItem} 作ったクリップグループ
     */
    function clipImageWithPath(sourceImage, maskPath, placeAnchor) {
        var anchorItem = placeAnchor || maskPath;
        var clipGroup = anchorItem.parent.groupItems.add();
        /* 置き場所のすぐ前面にグループを置いてから中へ入れ、重ね順を保つ / Put the group right in front of the anchor, then move items in */
        clipGroup.move(anchorItem, ElementPlacement.PLACEBEFORE);
        sourceImage.duplicate(clipGroup, ElementPlacement.PLACEATBEGINNING);
        maskPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;
        return clipGroup;
    }

    // =========================================
    // 線で区切る / Divide with lines
    // =========================================

    /* 領域が画像の範囲に収まっているかの許容値（pt） / Tolerance (pt) for the inside-the-image check */
    var REGION_BOUNDS_TOLERANCE = 0.01;

    /**
     * 境界 innerBounds が outerBounds の中に収まっているか
     * @param {number[]} innerBounds - 内側の境界 [左, 上, 右, 下]
     * @param {number[]} outerBounds - 外側の境界 [左, 上, 右, 下]
     * @returns {boolean} 収まっていれば true
     */
    function boundsInside(innerBounds, outerBounds) {
        var tolerance = REGION_BOUNDS_TOLERANCE;
        return innerBounds[0] >= outerBounds[0] - tolerance && innerBounds[1] <= outerBounds[1] + tolerance &&
            innerBounds[2] <= outerBounds[2] + tolerance && innerBounds[3] >= outerBounds[3] - tolerance;
    }

    /**
     * 分割で残った閉じた領域（パス／複合パス）を集める
     * @param {PageItem} item - 分割・拡張したアイテム
     * @param {number[]} imageBounds - 画像の境界
     * @param {Array} regions - 見つけた領域を入れる配列
     * @returns {void}
     */
    function collectRegions(item, imageBounds, regions) {
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) collectRegions(item.pageItems[i], imageBounds, regions);
            return;
        }
        var isRegion = item.typename === "CompoundPathItem" ||
            (item.typename === "PathItem" && item.closed && Math.abs(item.area) > 0);
        /* 画像の外にはみ出た領域は、区切り線の外側の切れ端なので除く / Pieces sticking out of the image come from outside the image area */
        if (isRegion && boundsInside(item.geometricBounds, imageBounds)) regions.push(item);
    }

    /**
     * 画像の範囲の長方形をパスで分割（パスファインダーの［分割］）し、できた領域を返す
     * @param {PlacedItem|RasterItem} sourceImage - 元の画像
     * @param {Array} cutterPaths - 区切りに使うパス
     * @returns {{ workItem: PageItem, regions: Array }} 分割・拡張したアイテムと、画像の範囲内の領域
     */
    function divideImageArea(sourceImage, cutterPaths) {
        var doc = app.activeDocument;
        var imageBounds = sourceImage.geometricBounds;
        /* 効果は入れ子のグループだと外側に付くので、レイヤー直下で作業する / Effects land on the outermost group, so work at layer level */
        var workGroup = sourceImage.layer.groupItems.add();
        for (var i = 0; i < cutterPaths.length; i++) {
            var cutterCopy = cutterPaths[i].duplicate(workGroup, ElementPlacement.PLACEATEND);
            if (!cutterCopy.filled && !cutterCopy.stroked) cutterCopy.stroked = true;
        }
        var areaRect = workGroup.pathItems.rectangle(imageBounds[1], imageBounds[0],
            imageBounds[2] - imageBounds[0], imageBounds[1] - imageBounds[3]);
        /* 最前面に置き、分割した領域がこの塗りを受け継ぐようにする / Keep it frontmost so the divided faces take its fill */
        areaRect.move(workGroup, ElementPlacement.PLACEATBEGINNING);
        var areaFill = new GrayColor();
        areaFill.gray = 50;
        areaRect.filled = true;
        areaRect.fillColor = areaFill;
        areaRect.stroked = false;

        doc.selection = [workGroup];
        app.redraw();
        app.executeMenuCommand("Live Pathfinder Divide");
        /* 効果を掛けた直後の分割・拡張は効かないので、選び直して再描画する / Reselect and redraw; expandStyle right after an effect does nothing */
        doc.selection = [workGroup];
        app.redraw();
        app.executeMenuCommand("expandStyle");

        var workItem = (doc.selection.length > 0) ? doc.selection[0] : workGroup;
        var regions = [];
        collectRegions(workItem, imageBounds, regions);
        doc.selection = null;
        return { workItem: workItem, regions: regions };
    }

    /**
     * 画像の範囲をパスで区切り、区切られた領域ごとにマスクする
     * @param {PlacedItem|RasterItem} sourceImage - 元の画像
     * @param {Array} cutterPaths - 区切りに使うパス
     * @returns {Array} 作ったクリップグループ（領域を作れなければ空）
     */
    function clipImageByDividedRegions(sourceImage, cutterPaths) {
        var divided = divideImageArea(sourceImage, cutterPaths);
        var clipGroups = [];
        /* 領域はまとめて元の画像のすぐ前面に置く / Every region goes right in front of the original image */
        for (var i = 0; i < divided.regions.length; i++) {
            clipGroups.push(clipImageWithPath(sourceImage, divided.regions[i], sourceImage));
        }
        divided.workItem.remove();
        return clipGroups;
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

        var hasDividerLine = false;
        for (var m = 0; m < overlappingPaths.length; m++) {
            if (isDividerLine(overlappingPaths[m])) hasDividerLine = true;
        }

        doc.selection = null;
        var clipGroups = [];
        if (hasDividerLine) {
            /* 線があれば、すべてのパスで画像の範囲を区切る / With a line in the selection, every path divides the image area */
            clipGroups = clipImageByDividedRegions(sourceImage, overlappingPaths);
            if (clipGroups.length === 0) {
                doc.selection = selectedItems;
                alert(getLabel(LABELS.alert.noRegions));
                return;
            }
            for (var n = 0; n < overlappingPaths.length; n++) overlappingPaths[n].remove();
        } else {
            for (var k = 0; k < overlappingPaths.length; k++) {
                clipGroups.push(clipImageWithPath(sourceImage, overlappingPaths[k]));
            }
        }
        sourceImage.remove();
        doc.selection = clipGroups;
    }

    main();

})();
