#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ClipImageWithEachPath
Name=画像をパスごとに切り分け
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.0
Release-Date=2026-11-03
Description-BEGIN
選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。
選択に画像を横切る線（始点と終点が離れたオープンパス）が含まれていると、画像の範囲をパスで区切り、区切られた領域ごとにマスクを作ります。

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.
When the selection includes lines across the image (open paths whose ends are far apart), the image area is divided by the paths and each resulting region becomes a mask.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択した画像を、一緒に選択したパスの数だけ複製し、それぞれのパスでクリッピングマスクを作ります。
1枚の画像に描き並べたイラストを、パスで囲んで別々のオブジェクトに切り分けるときに使います。
選択に画像を横切る線（始点と終点が離れたオープンパス）が含まれていると、画像の範囲をパスで区切り、区切られた領域ごとにマスクを作ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6b28a104130d

### Overview

Duplicates the selected image once for each selected path and clips each copy with its own path.
Use it to cut several illustrations drawn on one image into separate objects.
When the selection includes lines across the image (open paths whose ends are far apart), the image area is divided by the paths and each resulting region becomes a mask.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipImageWithEachPath";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-04";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "画像をパスごとに切り分け"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipImageWithEachPath.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipImageWithEachPath.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6b28a104130d"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipImageWithEachPath.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 区切り線の両端を延ばす長さ（pt）。画像の縁の手前で止まった線も縁まで届かせる / Length (pt) to extend both ends of a dividing line so lines that stop short of the edge still reach it */
    var DIVIDER_EXTEND_LENGTH = 3;

    /* 始点と終点の距離が、パスの縦横の長いほうのこの割合より大きいオープンパスを区切り線とみなす / An open path is a dividing line when its ends are farther apart than this share of its longer side */
    var DIVIDER_ENDPOINT_GAP_RATIO = 0.5;

    /* これより小さい領域（面積 pt²）はマスクにしない / Regions smaller than this area (pt²) are not clipped */
    var REGION_MIN_AREA = 1;

    /* 平均の幅（面積÷長いほうの辺、pt）がこれより細い領域はマスクにしない / Regions thinner than this average width (area / longer side, pt) are not clipped */
    var REGION_MIN_WIDTH = 0.5;

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
     * 選択から画像とパスを集める。グループ（クリップグループを除く）は中まで見る
     * @param {Array} items - 調べるアイテム
     * @param {Array} images - 見つけた画像を入れる配列
     * @param {Array} paths - 見つけたパスを入れる配列
     * @returns {void}
     */
    function collectImagesAndPaths(items, images, paths) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            /* グループの中のロック・非表示のアイテムは動かせないので使わない / Locked or hidden items inside groups cannot be moved */
            if (item.locked || item.hidden) continue;
            if (isImageItem(item)) images.push(item);
            else if (isMaskPath(item)) paths.push(item);
            else if (item.typename === "GroupItem" && !item.clipped) collectImagesAndPaths(item.pageItems, images, paths);
        }
    }

    /**
     * 区切り線として使う線か（閉じていないパスで、始点と終点が離れている）
     * 鉛筆ツールなどで囲んで閉じていないパスは、始点と終点がほぼ重なるので囲みとして扱う
     * @param {PathItem|CompoundPathItem} item - 判定するパス
     * @returns {boolean} 区切り線なら true
     */
    function isDividerLine(item) {
        if (item.typename !== "PathItem" || item.closed || item.pathPoints.length < 2) return false;
        var points = item.pathPoints;
        var firstAnchor = points[0].anchor;
        var lastAnchor = points[points.length - 1].anchor;
        var bounds = item.geometricBounds;
        var longerSide = Math.max(bounds[2] - bounds[0], bounds[1] - bounds[3]);
        var dx = lastAnchor[0] - firstAnchor[0];
        var dy = lastAnchor[1] - firstAnchor[1];
        return Math.sqrt(dx * dx + dy * dy) > longerSide * DIVIDER_ENDPOINT_GAP_RATIO;
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

    /* 方向線が付いていないとみなす距離（pt） / Distance (pt) under which a direction handle counts as absent */
    var HANDLE_EPSILON = 0.001;

    /**
     * 2点が（ほぼ）同じ位置か
     * @param {number[]} pointA - 1つめの点 [x, y]
     * @param {number[]} pointB - 2つめの点 [x, y]
     * @returns {boolean} 同じ位置なら true
     */
    function isSamePoint(pointA, pointB) {
        return Math.abs(pointA[0] - pointB[0]) < HANDLE_EPSILON && Math.abs(pointA[1] - pointB[1]) < HANDLE_EPSILON;
    }

    /**
     * 線の端点から、線の外へ向かう単位ベクトルを返す（端の方向線、なければ隣のアンカーの方向から求める）
     * @param {PathPoint} endPoint - 端のアンカーポイント
     * @param {number[]} endHandle - 端点の、線の内側へ向かう方向線
     * @param {PathPoint} neighborPoint - 隣のアンカーポイント
     * @param {number[]} neighborHandle - 隣のアンカーの、端点へ向かう方向線
     * @returns {number[]|null} 単位ベクトル [x, y]（求められなければ null）
     */
    function getOutwardDirection(endPoint, endHandle, neighborPoint, neighborHandle) {
        var endAnchor = endPoint.anchor;
        var innerPoint = neighborPoint.anchor;
        if (!isSamePoint(endHandle, endAnchor)) innerPoint = endHandle;
        else if (!isSamePoint(neighborHandle, endAnchor)) innerPoint = neighborHandle;
        var dx = endAnchor[0] - innerPoint[0];
        var dy = endAnchor[1] - innerPoint[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length < HANDLE_EPSILON) return null;
        return [dx / length, dy / length];
    }

    /**
     * パスの末尾にアンカーポイントを足す
     * @param {PathItem} pathItem - 足す先のパス
     * @param {number[]} anchor - アンカーの位置
     * @param {number[]} leftDirection - 前側の方向線
     * @param {number[]} rightDirection - 後ろ側の方向線
     * @param {PointType} pointType - ポイントの種類
     * @returns {void}
     */
    function appendPathPoint(pathItem, anchor, leftDirection, rightDirection, pointType) {
        var newPoint = pathItem.pathPoints.add();
        newPoint.pointType = pointType;
        newPoint.anchor = anchor;
        newPoint.leftDirection = leftDirection;
        newPoint.rightDirection = rightDirection;
    }

    /**
     * 区切り線を、両端を DIVIDER_EXTEND_LENGTH だけ延ばした形で container に描き直す
     * @param {PathItem} linePath - 区切り線
     * @param {GroupItem} container - 描く先のグループ
     * @returns {PathItem} 延ばした線
     */
    function addExtendedLineCopy(linePath, container) {
        var sourcePoints = linePath.pathPoints;
        var lastIndex = sourcePoints.length - 1;
        var firstPoint = sourcePoints[0];
        var lastPoint = sourcePoints[lastIndex];
        var startDirection = getOutwardDirection(firstPoint, firstPoint.rightDirection, sourcePoints[1], sourcePoints[1].leftDirection);
        var endDirection = getOutwardDirection(lastPoint, lastPoint.leftDirection, sourcePoints[lastIndex - 1], sourcePoints[lastIndex - 1].rightDirection);

        var extendedLine = container.pathItems.add();
        if (startDirection) {
            var startTip = [firstPoint.anchor[0] + startDirection[0] * DIVIDER_EXTEND_LENGTH,
                firstPoint.anchor[1] + startDirection[1] * DIVIDER_EXTEND_LENGTH];
            appendPathPoint(extendedLine, startTip, startTip, startTip, PointType.CORNER);
        }
        for (var i = 0; i <= lastIndex; i++) {
            var sourcePoint = sourcePoints[i];
            appendPathPoint(extendedLine, sourcePoint.anchor, sourcePoint.leftDirection, sourcePoint.rightDirection, sourcePoint.pointType);
        }
        if (endDirection) {
            var endTip = [lastPoint.anchor[0] + endDirection[0] * DIVIDER_EXTEND_LENGTH,
                lastPoint.anchor[1] + endDirection[1] * DIVIDER_EXTEND_LENGTH];
            appendPathPoint(extendedLine, endTip, endTip, endTip, PointType.CORNER);
        }
        extendedLine.closed = false;
        extendedLine.filled = false;
        extendedLine.stroked = true;
        return extendedLine;
    }

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
     * 領域の面積を返す（複合パスはいちばん大きいサブパスの面積）
     * @param {PathItem|CompoundPathItem} region - 領域
     * @returns {number} 面積（pt²）
     */
    function getRegionArea(region) {
        if (region.typename === "PathItem") return Math.abs(region.area);
        var largestArea = 0;
        for (var i = 0; i < region.pathItems.length; i++) {
            largestArea = Math.max(largestArea, Math.abs(region.pathItems[i].area));
        }
        return largestArea;
    }

    /**
     * 線どうしのわずかな交差などでできた、小さすぎる・細すぎる領域か
     * @param {PathItem|CompoundPathItem} region - 領域
     * @returns {boolean} マスクにしないなら true
     */
    function isTooSmallRegion(region) {
        var area = getRegionArea(region);
        if (area < REGION_MIN_AREA) return true;
        var bounds = region.geometricBounds;
        var longerSide = Math.max(bounds[2] - bounds[0], bounds[1] - bounds[3]);
        return longerSide > 0 && area / longerSide < REGION_MIN_WIDTH;
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
        var isRegion = item.typename === "CompoundPathItem" || (item.typename === "PathItem" && item.closed);
        if (!isRegion || isTooSmallRegion(item)) return;
        /* 画像の外にはみ出た領域は、区切り線の外側の切れ端なので除く / Pieces sticking out of the image come from outside the image area */
        if (boundsInside(item.geometricBounds, imageBounds)) regions.push(item);
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
            if (isDividerLine(cutterPaths[i])) {
                addExtendedLineCopy(cutterPaths[i], workGroup);
                continue;
            }
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

    /* 同じ列の領域の上端がそろっているとみなす差（pt） / Difference (pt) under which tops in one column count as level */
    var SORT_LEVEL_TOLERANCE = 0.5;

    /**
     * 領域 A を領域 B より先に並べるか（左から右、同じ列なら上から下）
     * @param {number[]} boundsA - 領域 A の境界
     * @param {number[]} boundsB - 領域 B の境界
     * @returns {boolean} A が先なら true
     */
    function comesBefore(boundsA, boundsB) {
        var centerA = (boundsA[0] + boundsA[2]) / 2;
        var centerB = (boundsB[0] + boundsB[2]) / 2;
        /* どちらかの中心がもう一方の左右の範囲に入っていれば同じ列 / Same column when either center falls within the other's width */
        var sameColumn = (centerA > boundsB[0] && centerA < boundsB[2]) || (centerB > boundsA[0] && centerB < boundsA[2]);
        if (sameColumn && Math.abs(boundsA[1] - boundsB[1]) > SORT_LEVEL_TOLERANCE) return boundsA[1] > boundsB[1];
        return centerA < centerB;
    }

    /**
     * 領域を左から右（同じ列なら上から下）に並べ替える（比較関数つき sort() は遅く並びも狂うので挿入ソート）
     * @param {Array} regions - 領域
     * @returns {Array} 並べ替えた領域
     */
    function sortRegionsLeftToRight(regions) {
        var entries = [];
        for (var i = 0; i < regions.length; i++) entries.push({ item: regions[i], bounds: regions[i].geometricBounds });
        for (var j = 1; j < entries.length; j++) {
            var current = entries[j];
            var k = j - 1;
            while (k >= 0 && comesBefore(current.bounds, entries[k].bounds)) {
                entries[k + 1] = entries[k];
                k--;
            }
            entries[k + 1] = current;
        }
        var sorted = [];
        for (var m = 0; m < entries.length; m++) sorted.push(entries[m].item);
        return sorted;
    }

    /**
     * 画像の範囲をパスで区切り、区切られた領域ごとにマスクする
     * @param {PlacedItem|RasterItem} sourceImage - 元の画像
     * @param {Array} cutterPaths - 区切りに使うパス
     * @returns {Array} 作ったクリップグループ（領域を作れなければ空）
     */
    function clipImageByDividedRegions(sourceImage, cutterPaths) {
        var divided = divideImageArea(sourceImage, cutterPaths);
        var regions = sortRegionsLeftToRight(divided.regions);
        var clipGroups = [];
        /* 元の画像のすぐ前面に順に置くと、先に作ったものほど上に並ぶ / Placing each right in front of the image keeps earlier ones on top */
        for (var i = 0; i < regions.length; i++) {
            clipGroups.push(clipImageWithPath(sourceImage, regions[i], sourceImage));
        }
        divided.workItem.remove();
        return clipGroups;
    }

    /**
     * 配列にアイテムが入っているか（DOM の参照は === で同一判定できる）
     * @param {Array} items - 配列
     * @param {PageItem} target - 探すアイテム
     * @returns {boolean} 入っていれば true
     */
    function containsItem(items, target) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === target) return true;
        }
        return false;
    }

    /**
     * パスを削除し、それで空になったグループも消す
     * @param {Array} paths - 削除するパス
     * @returns {void}
     */
    function removePathsAndEmptyGroups(paths) {
        var parentGroups = [];
        for (var i = 0; i < paths.length; i++) {
            var parentItem = paths[i].parent;
            if (parentItem.typename === "GroupItem" && !containsItem(parentGroups, parentItem)) parentGroups.push(parentItem);
            paths[i].remove();
        }
        /* 内側のグループが空になって消えると外側も空になることがあるので、親へさかのぼる / Removing an empty inner group can empty its parent, so walk upward */
        for (var j = 0; j < parentGroups.length; j++) {
            var group = parentGroups[j];
            while (group.typename === "GroupItem" && group.pageItems.length === 0) {
                var outerItem = group.parent;
                group.remove();
                group = outerItem;
                /* ほかの子の親として控えたグループを二度消さない / Do not remove a group recorded for another path twice */
                for (var k = j + 1; k < parentGroups.length; k++) {
                    if (parentGroups[k] === group) parentGroups.splice(k, 1);
                }
            }
        }
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
        collectImagesAndPaths(selectedItems, selectedImages, maskPaths);
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
            /* 区切り線があれば、すべてのパスで画像の範囲を区切る / With a dividing line in the selection, every path divides the image area */
            clipGroups = clipImageByDividedRegions(sourceImage, overlappingPaths);
            if (clipGroups.length === 0) {
                doc.selection = selectedItems;
                alert(getLabel(LABELS.alert.noRegions));
                return;
            }
            removePathsAndEmptyGroups(overlappingPaths);
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
