#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを高さ・幅・不透明度・カラーなどの基準で並び替え、横または縦に整列・分布させます。
整列方向、基準、順序、間隔、幅・高さの統一を、プレビューを見ながら指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectSorter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n663264db75ff

### Overview

Sorts the selected objects by height, width, opacity or color and then aligns and distributes them horizontally or vertically.
Direction, sort key, order, spacing and size unification are all set while watching a preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectSorter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectSorter";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectSorter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectSorter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n663264db75ff"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「指定」間隔の初期値（pt） / Initial custom gap in points */
    var DEFAULT_CUSTOM_GAP = "20";

    /* 「数字」で並べるときの上端の送り量（pt） / Top-to-top step for the Number sort, in points */
    var NUMBER_SORT_STEP = 50;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの整列", en: "Object Alignment Tool" }
        },
        panel: {
            sortKey: { ja: "基準", en: "Sort by" },
            sortOrder: { ja: "ソート順", en: "Sort Order" },
            vertical: { ja: "縦方向", en: "Vertical" },
            alignVertical: { ja: "揃え", en: "Align Vertically" },
            spacingVertical: { ja: "縦間隔", en: "Vertical Spacing" },
            matchWidth: { ja: "幅を揃える", en: "Match Width" },
            horizontal: { ja: "横方向", en: "Horizontal" },
            alignHorizontal: { ja: "揃え", en: "Align Horizontally" },
            spacingHorizontal: { ja: "横間隔", en: "Horizontal Spacing" },
            matchHeight: { ja: "高さを揃える", en: "Match Height" }
        },
        radio: {
            alongX: { ja: "横並び", en: "Horizontal" },
            alongY: { ja: "縦並び", en: "Vertical" },
            byHeight: { ja: "高さ", en: "Height" },
            byWidth: { ja: "幅", en: "Width" },
            byOpacity: { ja: "不透明度", en: "Opacity" },
            byColor: { ja: "カラー", en: "Color" },
            byNumber: { ja: "数字", en: "Number" },
            byZOrder: { ja: "重ね順", en: "Z-Order" },
            ascending: { ja: "昇順", en: "Ascending" },
            descending: { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" },
            spacingEven: { ja: "均等", en: "Even" },
            spacingTight: { ja: "ぴったり", en: "Tight" },
            spacingCustom: { ja: "指定", en: "Custom" },
            matchMax: { ja: "最大", en: "Max" },
            matchMin: { ja: "最小", en: "Min" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        tooltip: {
            alongX: {
                ja: "基準とソート順に従って、オブジェクトどうしの横位置を入れ替えます",
                en: "Swaps the objects' horizontal positions to follow the sort key and order"
            },
            alongY: {
                ja: "基準とソート順に従って、オブジェクトどうしの縦位置を入れ替えます",
                en: "Swaps the objects' vertical positions to follow the sort key and order"
            },
            byColor: {
                ja: "塗りのカラーをグレースケールに換算した値で並べ替えます",
                en: "Sorts by the fill color converted to a grayscale value"
            },
            byNumber: {
                ja: "数字だけのテキストを含むグループを数値の順に並べ、上端を50 ptずつずらして縦に配置します",
                en: "Orders groups that contain a digits-only text by that number and stacks them, each top 50 pt below the previous one"
            },
            spacingEven: {
                ja: "両端のオブジェクトの位置はそのままで、間隔を均等にします",
                en: "Evens out the gaps while the objects at both ends stay put"
            },
            spacingTight: { ja: "間隔を0にして詰めます", en: "Closes the gaps between objects" },
            customGap: { ja: "オブジェクトの間隔（pt）", en: "Gap between objects, in points" },
            matchMaxWidth: { ja: "いちばん広い幅に合わせて、幅だけを変えます", en: "Scales only the width to match the widest object" },
            matchMinWidth: { ja: "いちばん狭い幅に合わせて、幅だけを変えます", en: "Scales only the width to match the narrowest object" },
            matchMaxHeight: { ja: "いちばん高い高さに合わせて、高さだけを変えます", en: "Scales only the height to match the tallest object" },
            matchMinHeight: { ja: "いちばん低い高さに合わせて、高さだけを変えます", en: "Scales only the height to match the shortest object" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "実行", en: "Apply" }
        },
        alert: {
            noSelection: { ja: "ファイルを開き、並べ替えるオブジェクトを選択してください。", en: "Open a file and select objects to sort." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなパス
     * @returns {string} 表示言語のテキスト（見つからなければパスそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en;
    }

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /* 基準・並び方向のキーと、比べる／入れ替えるプロパティ名の対応
       Option key -> property that is compared or reassigned */
    var OPTION_PROPERTIES = {
        h: "height",
        w: "width",
        o: "opacity",
        color: "color",
        z: "zOrderPosition",
        x: "left",
        y: "top"
    };

    /**
     * 数値を昇順に比べる
     * @param {number} a - 比べる値
     * @param {number} b - 比べる値
     * @returns {number} 並べ替え用の差
     */
    function compareAscending(a, b) {
        return a - b;
    }

    /**
     * 数値を降順に比べる
     * @param {number} a - 比べる値
     * @param {number} b - 比べる値
     * @returns {number} 並べ替え用の差
     */
    function compareDescending(a, b) {
        return b - a;
    }

    /**
     * 並べ替えの順番をランダムにする
     * @returns {number} -0.5〜0.5 の乱数
     */
    function compareRandomly() {
        return Math.random() - .5;
    }

    /**
     * カラーを色成分の配列にする
     * @param {Color} sourceColor - 対象のカラー
     * @returns {number[]} CMYK は4つ、RGB は3つ、グレーは1つの成分（その他は [0]）
     */
    function getColorChannels(sourceColor) {
        var color = sourceColor;
        if (color.hasOwnProperty('color')) color = color.color;
        if (color.constructor.name == 'SpotColor') color = color.spot.color;

        if (color.constructor.name === 'CMYKColor') return [color.cyan, color.magenta, color.yellow, color.black];
        if (color.constructor.name === 'RGBColor') return [color.red, color.green, color.blue];
        if (color.constructor.name === 'GrayColor') return [color.gray];
        return [0];
    }

    /**
     * カラーをグレースケールに換算する
     * @param {Color} sourceColor - 対象のカラー
     * @returns {number[]} グレースケールの値（要素1つの配列）
     */
    function getGrayScaleValue(sourceColor) {
        var channels = getColorChannels(sourceColor);
        if (channels.length === 4) {
            return app.convertSampleColor(ImageColorSpace.CMYK, channels, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
        }
        if (channels.length === 3) {
            return app.convertSampleColor(ImageColorSpace.RGB, channels, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
        }
        return channels;
    }

    /**
     * 「カラー」基準で比べる値（塗りのグレースケール値）を返す
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]|number} グレースケールの値。塗りが無ければ 0
     */
    function getColorSortValue(item) {
        return item.fillColor ? getGrayScaleValue(item.fillColor) : 0;
    }

    /**
     * 指定プロパティでオブジェクトを比べる関数を作る
     * @param {string} propertyName - 比べるプロパティ名（"color" は塗りの明るさ）
     * @returns {Function} Array.sort 用の比較関数
     */
    function createItemComparator(propertyName) {
        return function (a, b) {
            if (propertyName === "color") {
                return getColorSortValue(a) - getColorSortValue(b);
            }
            return Number(a[propertyName]) - Number(b[propertyName]);
        };
    }

    /**
     * オブジェクトを基準の順に並べ、その順に位置の値を割り当て直す
     * @param {PageItem[]} items - 対象オブジェクト（この配列自体も並べ替える）
     * @param {Function} compareItems - オブジェクトの比較関数
     * @param {Function} comparePositions - 位置の値の比較関数（昇順／降順／ランダム）
     * @param {string} positionProperty - 割り当て直すプロパティ（"left" / "top"）
     * @returns {void}
     */
    function rearrangeItems(items, compareItems, comparePositions, positionProperty) {
        items.sort(compareItems);
        var positions = [];
        for (var i = 0; i < items.length; i++) {
            positions.push(items[i][positionProperty]);
        }
        positions.sort(comparePositions);
        /* 上端は値が大きいほど上なので逆順に / larger top means higher, so reverse */
        if (positionProperty === "top") positions.reverse();
        for (var j = 0; j < items.length; j++) {
            items[j][positionProperty] = positions[j];
        }
    }

    /**
     * グループ内で最初に見つかった数字だけのテキストを数値で返す（入れ子のグループも探す）
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {number} 見つかった数値。無ければ NaN
     */
    function findNumberInGroup(groupItem) {
        var memberItems = groupItem.pageItems;
        for (var i = 0; i < memberItems.length; i++) {
            var memberItem = memberItems[i];
            if (memberItem.typename === "TextFrame") {
                var frameText = memberItem.contents;
                if (/^\d+$/.test(frameText)) {
                    return parseFloat(frameText);
                }
            } else if (memberItem.typename === "GroupItem") {
                var nestedNumber = findNumberInGroup(memberItem);
                if (!isNaN(nestedNumber)) return nestedNumber;
            }
        }
        return NaN;
    }

    /**
     * 数字を含むグループを数値の順に並べ、先頭のグループの位置から縦に配置する
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} orderKey - ソート順のキー（"l" は降順、それ以外は昇順）
     * @returns {void}
     */
    function arrangeGroupsByNumber(items, orderKey) {
        var numberedGroups = [];
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "GroupItem") {
                var groupNumber = findNumberInGroup(items[i]);
                if (!isNaN(groupNumber)) {
                    numberedGroups.push({ group: items[i], value: groupNumber });
                }
            }
        }
        if (numberedGroups.length === 0) return;
        if (orderKey === "l") {
            numberedGroups.sort(function (a, b) { return b.value - a.value; });
        } else {
            numberedGroups.sort(function (a, b) { return a.value - b.value; });
        }
        var startTop = numberedGroups[0].group.top;
        var startLeft = numberedGroups[0].group.left;
        for (var j = 0; j < numberedGroups.length; j++) {
            numberedGroups[j].group.left = startLeft;
            numberedGroups[j].group.top = startTop - j * NUMBER_SORT_STEP;
        }
    }

    /**
     * 基準・並び方向・ソート順に従ってオブジェクトを並べ替える
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} sortKey - 基準のキー（h / w / o / color / n / z）
     * @param {string} alongKey - 並び方向のキー（x / y）
     * @param {string} orderKey - ソート順のキー（s / l / r）
     * @returns {void}
     */
    function applyArrangement(items, sortKey, alongKey, orderKey) {
        if (sortKey === "n") {
            arrangeGroupsByNumber(items, orderKey);
            return;
        }
        var comparePositions = compareAscending;
        if (orderKey === "l") {
            comparePositions = compareDescending;
        } else if (orderKey === "r") {
            comparePositions = compareRandomly;
        }
        rearrangeItems(items, createItemComparator(OPTION_PROPERTIES[sortKey]), comparePositions, OPTION_PROPERTIES[alongKey]);
    }

    // =========================================
    // 間隔と揃え / Spacing and alignment
    // =========================================

    /**
     * 両端の位置を保ったまま均等にするときの間隔を求める
     * @param {PageItem[]} sortedItems - 並び順に並べたオブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {number} オブジェクト間の間隔（pt）
     */
    function getEvenGap(sortedItems, isHorizontal) {
        var firstItem = sortedItems[0];
        var lastItem = sortedItems[sortedItems.length - 1];
        var totalSize = 0;
        for (var i = 0; i < sortedItems.length; i++) {
            totalSize += isHorizontal ? sortedItems[i].width : sortedItems[i].height;
        }
        var totalGap = isHorizontal
            ? (lastItem.left + lastItem.width) - firstItem.left - totalSize
            : firstItem.top - (lastItem.top - lastItem.height) - totalSize;
        return totalGap / (sortedItems.length - 1);
    }

    /**
     * 先頭のオブジェクトの位置から、指定の間隔で順に並べる
     * @param {PageItem[]} sortedItems - 並び順に並べたオブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @param {number} gap - オブジェクト間の間隔（pt）
     * @returns {void}
     */
    function placeWithGap(sortedItems, isHorizontal, gap) {
        var position = isHorizontal ? sortedItems[0].left : sortedItems[0].top;
        for (var i = 0; i < sortedItems.length; i++) {
            if (isHorizontal) {
                sortedItems[i].left = position;
                position += sortedItems[i].width + gap;
            } else {
                sortedItems[i].top = position;
                position -= sortedItems[i].height + gap;
            }
        }
    }

    /**
     * 間隔の種類に従ってオブジェクトを並べ直す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} spacingType - "even"（均等）/ "zero"（ぴったり）/ "custom"（指定）
     * @param {EditText} gapInput - 「指定」の間隔の入力欄
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {void}
     */
    function applySpacing(items, spacingType, gapInput, isHorizontal) {
        var sortedItems = items.slice();
        sortedItems.sort(function (a, b) {
            return isHorizontal ? a.left - b.left : b.top - a.top;
        });
        var gap = 0;
        if (spacingType === "even") {
            gap = getEvenGap(sortedItems, isHorizontal);
        } else if (spacingType === "custom" && gapInput.text !== "") {
            gap = parseFloat(gapInput.text);
        }
        placeWithGap(sortedItems, isHorizontal, gap);
    }

    /**
     * 揃えの基準にする位置を、オブジェクトの左端または上端からのずれで返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} alignType - "left" / "center" / "right" / "top" / "middle" / "bottom"
     * @returns {number} 左端（上端）から基準位置までのずれ
     */
    function getAlignOffset(bounds, alignType) {
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        if (alignType === "center") return width / 2;
        if (alignType === "right") return width;
        if (alignType === "middle") return -(height / 2);
        if (alignType === "bottom") return -height;
        return 0;
    }

    /**
     * オブジェクトの左端・中央・右端（上端・中央・下端）を揃える
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} alignType - "left" / "center" / "right" / "top" / "middle" / "bottom"
     * @returns {void}
     */
    function alignItems(items, alignType) {
        if (!items || items.length === 0) return;
        var movesVertically = (alignType === "top" || alignType === "middle" || alignType === "bottom");

        /* 注意: previewBoundsCheckbox はダイアログ内のローカル変数で、ここからは見えないため常に false
           Note: previewBoundsCheckbox is local to the dialog, so this is always false here */
        var usePreviewBounds = (typeof previewBoundsCheckbox !== "undefined" && previewBoundsCheckbox.value);

        var itemBounds = [];
        var alignPositions = [];
        for (var i = 0; i < items.length; i++) {
            itemBounds[i] = usePreviewBounds ? items[i].visibleBounds : items[i].geometricBounds;
            var edgePosition = movesVertically ? itemBounds[i][1] : itemBounds[i][0];
            alignPositions.push(edgePosition + getAlignOffset(itemBounds[i], alignType));
        }

        var targetPosition;
        if (alignType === "top" || alignType === "left") {
            targetPosition = Math.min.apply(null, alignPositions);
        } else if (alignType === "bottom" || alignType === "right") {
            targetPosition = Math.max.apply(null, alignPositions);
        } else {
            var positionSum = 0;
            for (var j = 0; j < alignPositions.length; j++) {
                positionSum += alignPositions[j];
            }
            targetPosition = positionSum / alignPositions.length;
        }

        for (var k = 0; k < items.length; k++) {
            var alignOffset = getAlignOffset(itemBounds[k], alignType);
            if (movesVertically) {
                var dy = items[k].top - itemBounds[k][1];
                items[k].top = targetPosition - alignOffset + dy;
            } else {
                var dx = items[k].left - itemBounds[k][0];
                items[k].left = targetPosition - alignOffset + dx;
            }
        }
    }

    /**
     * オブジェクトの幅または高さを、いちばん大きい（小さい）ものに合わせて拡大・縮小する
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} sizeProperty - "width" / "height"
     * @param {boolean} useLargest - true なら最大、false なら最小に合わせる
     * @returns {void}
     */
    function matchItemSize(items, sizeProperty, useLargest) {
        var targetSize = useLargest ? 0 : Infinity;
        for (var i = 0; i < items.length; i++) {
            var itemSize = items[i][sizeProperty];
            if (useLargest ? itemSize > targetSize : itemSize < targetSize) {
                targetSize = itemSize;
            }
        }
        for (var j = 0; j < items.length; j++) {
            var scalePercent = targetSize / items[j][sizeProperty] * 100;
            if (sizeProperty === "width") {
                items[j].resize(scalePercent, 100);
            } else {
                items[j].resize(100, scalePercent);
            }
        }
        app.redraw();
    }

    // =========================================
    // 選択と位置 / Selection and positions
    // =========================================

    /**
     * 先頭3つのオブジェクトの中心から、横並びか縦並びかを推定する
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {string} 縦並びなら "y"、それ以外は "x"
     */
    function detectAlongDirection(items) {
        if (!(items.length >= 3)) return "x";
        var centers = [];
        for (var i = 0; i < 3; i++) {
            var visibleBounds = items[i].visibleBounds;
            centers.push({ x: (visibleBounds[0] + visibleBounds[2]) / 2, y: (visibleBounds[1] + visibleBounds[3]) / 2 });
        }
        var averageDx = (Math.abs(centers[0].x - centers[1].x) + Math.abs(centers[1].x - centers[2].x)) / 2;
        var averageDy = (Math.abs(centers[0].y - centers[1].y) + Math.abs(centers[1].y - centers[2].y)) / 2;
        return (averageDx * 1.5 < averageDy) ? "y" : "x";
    }

    /**
     * キャンセル時に戻せるよう、オブジェクトの位置を控える
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {Object[]} { item, left, top } の配列
     */
    function saveItemPositions(items) {
        var savedPositions = [];
        for (var i = 0; i < items.length; i++) {
            savedPositions.push({ item: items[i], left: items[i].left, top: items[i].top });
        }
        return savedPositions;
    }

    /**
     * 控えておいた位置へオブジェクトを戻す
     * @param {Object[]} savedPositions - saveItemPositions() の戻り値
     * @returns {void}
     */
    function restoreItemPositions(savedPositions) {
        for (var i = 0; i < savedPositions.length; i++) {
            savedPositions[i].item.left = savedPositions[i].left;
            savedPositions[i].item.top = savedPositions[i].top;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 選択肢のラジオボタンを並べて作る（値は optionKey に持たせる）
     * @param {Object} parentGroup - 追加先のグループまたはパネル
     * @param {Object[]} choices - { key, label, tooltip } の配列（label と tooltip は LABELS のパス）
     * @param {string} selectedKey - 最初に選んでおくキー
     * @returns {RadioButton[]} 作ったラジオボタン
     */
    function addChoiceRadios(parentGroup, choices, selectedKey) {
        var choiceRadios = [];
        for (var i = 0; i < choices.length; i++) {
            var choiceRadio = parentGroup.add("radiobutton", undefined, getLabel(choices[i].label));
            choiceRadio.optionKey = choices[i].key;
            choiceRadio.value = (choices[i].key === selectedKey);
            if (choices[i].tooltip) choiceRadio.helpTip = getLabel(choices[i].tooltip);
            choiceRadios.push(choiceRadio);
        }
        return choiceRadios;
    }

    /**
     * 選ばれているラジオボタンのキーを返す
     * @param {RadioButton[]} choiceRadios - ラジオボタン
     * @returns {string|null} optionKey。どれも選ばれていなければ null
     */
    function getSelectedKey(choiceRadios) {
        for (var i = 0; i < choiceRadios.length; i++) {
            if (choiceRadios[i].value) return choiceRadios[i].optionKey;
        }
        return null;
    }

    /**
     * 基準・ソート順のパネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {number} childSpacing - 項目の間隔
     * @returns {Panel} 作ったパネル
     */
    function addOptionPanel(parentGroup, titlePath, childSpacing) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = "left";
        optionPanel.spacing = childSpacing;
        optionPanel.margins = [15, 20, 15, 10];
        return optionPanel;
    }

    /**
     * 選択肢を横一列に並べるパネルを作る（揃え・幅／高さを揃える）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @returns {Panel} 作ったパネル
     */
    function addRowPanel(parentPanel, titlePath) {
        var rowPanel = parentPanel.add("panel", undefined, getLabel(titlePath));
        rowPanel.orientation = "row";
        rowPanel.alignChildren = "center";
        rowPanel.margins = [10, 20, 10, 10];
        return rowPanel;
    }

    /**
     * 揃えのパネルを作る。ラジオボタンを押すとすぐに揃える
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {Object[]} choices - { key, label } の配列（key は alignItems() の揃え方）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {void}
     */
    function addAlignPanel(parentPanel, titlePath, choices, targetItems) {
        var alignRadios = addChoiceRadios(addRowPanel(parentPanel, titlePath), choices, null);
        for (var i = 0; i < alignRadios.length; i++) {
            alignRadios[i].onClick = function () {
                alignItems(targetItems, this.optionKey);
                app.redraw();
            };
        }
    }

    /**
     * 幅（高さ）を揃えるパネルを作る。ラジオボタンを押すとすぐに拡大・縮小する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} sizeProperty - "width" / "height"
     * @returns {void}
     */
    function addMatchSizePanel(parentPanel, titlePath, targetItems, sizeProperty) {
        var matchesWidth = (sizeProperty === "width");
        var matchRadios = addChoiceRadios(addRowPanel(parentPanel, titlePath), [
            { key: "max", label: "radio.matchMax", tooltip: matchesWidth ? "tooltip.matchMaxWidth" : "tooltip.matchMaxHeight" },
            { key: "min", label: "radio.matchMin", tooltip: matchesWidth ? "tooltip.matchMinWidth" : "tooltip.matchMinHeight" }
        ], null);
        matchRadios[0].onClick = function () {
            matchItemSize(targetItems, sizeProperty, true);
        };
        matchRadios[1].onClick = function () {
            matchItemSize(targetItems, sizeProperty, false);
        };
    }

    /**
     * 間隔のパネル（均等／ぴったり／指定）を作る
     * 「指定」だけ別のグループに入るので、ラジオボタンの排他は手で管理する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {void}
     */
    function addSpacingPanel(parentPanel, titlePath, targetItems, isHorizontal) {
        var spacingPanel = parentPanel.add("panel", undefined, getLabel(titlePath));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = "left";
        spacingPanel.margins = [10, 20, 10, 10];

        var spacingRadioGroup = spacingPanel.add("group");
        spacingRadioGroup.orientation = "column";
        spacingRadioGroup.alignChildren = "left";
        var presetRadios = addChoiceRadios(spacingRadioGroup, [
            { key: "even", label: "radio.spacingEven", tooltip: "tooltip.spacingEven" },
            { key: "zero", label: "radio.spacingTight", tooltip: "tooltip.spacingTight" }
        ], null);

        var customGapGroup = spacingRadioGroup.add("group");
        customGapGroup.orientation = "row";
        customGapGroup.alignChildren = "left";
        var customRadios = addChoiceRadios(customGapGroup, [{ key: "custom", label: "radio.spacingCustom" }], null);
        var customGapInput = customGapGroup.add("edittext", undefined, DEFAULT_CUSTOM_GAP);
        customGapInput.characters = 5;
        customGapInput.enabled = false;
        customGapInput.helpTip = getLabel("tooltip.customGap");

        var spacingRadios = presetRadios.concat(customRadios);
        for (var i = 0; i < spacingRadios.length; i++) {
            spacingRadios[i].onClick = function () {
                var spacingType = this.optionKey;
                for (var j = 0; j < spacingRadios.length; j++) {
                    spacingRadios[j].value = (spacingRadios[j].optionKey === spacingType);
                }
                customGapInput.enabled = (spacingType === "custom");
                applySpacing(targetItems, spacingType, customGapInput, isHorizontal);
                app.redraw();
            };
        }
        customGapInput.onChange = function () {
            if (customRadios[0].value) {
                applySpacing(targetItems, "custom", customGapInput, isHorizontal);
                app.redraw();
            }
        };
    }

    /**
     * 縦方向・横方向のパネル（揃え・間隔・幅／高さを揃える）を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向のパネルなら true
     * @returns {void}
     */
    function addDirectionPanel(parentGroup, targetItems, isHorizontal) {
        var directionPanel = parentGroup.add("panel", undefined, getLabel(isHorizontal ? "panel.horizontal" : "panel.vertical"));
        directionPanel.alignChildren = "fill";
        directionPanel.margins = [10, 20, 10, 10];
        if (isHorizontal) {
            addAlignPanel(directionPanel, "panel.alignHorizontal", [
                { key: "top", label: "radio.alignTop" },
                { key: "middle", label: "radio.alignMiddle" },
                { key: "bottom", label: "radio.alignBottom" }
            ], targetItems);
            addSpacingPanel(directionPanel, "panel.spacingHorizontal", targetItems, true);
            addMatchSizePanel(directionPanel, "panel.matchHeight", targetItems, "height");
        } else {
            addAlignPanel(directionPanel, "panel.alignVertical", [
                { key: "left", label: "radio.alignLeft" },
                { key: "center", label: "radio.alignCenter" },
                { key: "right", label: "radio.alignRight" }
            ], targetItems);
            addSpacingPanel(directionPanel, "panel.spacingVertical", targetItems, false);
            addMatchSizePanel(directionPanel, "panel.matchWidth", targetItems, "width");
        }
    }

    /**
     * 並べ替え・整列のダイアログを開く。操作はその場でオブジェクトに反映し、キャンセルで位置を戻す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} defaultAlong - 最初に選んでおく並び方向（"x" / "y"）
     * @param {Object[]} savedPositions - キャンセル時に戻す位置
     * @returns {void}
     */
    function showSortDialog(targetItems, defaultAlong, savedPositions) {
        /* 縦並びなら基準を「幅」、横並びなら「高さ」から始める / start from Width when stacked vertically, Height otherwise */
        var defaultSortKey = (defaultAlong === "y") ? "w" : "h";

        var sortDialog = new Window("dialog", getLabel("dialog.title"));
        sortDialog.alignChildren = "left";
        sortDialog.orientation = "column";

        var alongGroup = sortDialog.add("group");
        alongGroup.alignment = "center";
        alongGroup.orientation = "row";
        alongGroup.spacing = 5;
        alongGroup.margins = [10, 10, 10, 10];
        var alongRadios = addChoiceRadios(alongGroup, [
            { key: "x", label: "radio.alongX", tooltip: "tooltip.alongX" },
            { key: "y", label: "radio.alongY", tooltip: "tooltip.alongY" }
        ], defaultAlong);

        var columnsGroup = sortDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = 10;

        var sortColumn = columnsGroup.add("group");
        sortColumn.orientation = "column";
        sortColumn.alignChildren = "fill";
        sortColumn.spacing = 10;

        var sortKeyRadios = addChoiceRadios(addOptionPanel(sortColumn, "panel.sortKey", 10), [
            { key: "h", label: "radio.byHeight" },
            { key: "w", label: "radio.byWidth" },
            { key: "o", label: "radio.byOpacity" },
            { key: "color", label: "radio.byColor", tooltip: "tooltip.byColor" },
            { key: "n", label: "radio.byNumber", tooltip: "tooltip.byNumber" },
            { key: "z", label: "radio.byZOrder" }
        ], defaultSortKey);

        var sortOrderRadios = addChoiceRadios(addOptionPanel(sortColumn, "panel.sortOrder", 5), [
            { key: "s", label: "radio.ascending" },
            { key: "l", label: "radio.descending" },
            { key: "r", label: "radio.random" }
        ], "s");

        addDirectionPanel(columnsGroup, targetItems, false);
        addDirectionPanel(columnsGroup, targetItems, true);

        var btnRowGroup = sortDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.margins = [10, 10, 10, 10];

        var previewBoundsCheckbox = btnRowGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = false;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 100;
        spacer.maximumSize.height = 0;

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOk = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOk.active = true;

        /* 並び方向・基準・ソート順を変えたらすぐに並べ替える / rearrange as soon as an option changes */
        var arrangeRadios = alongRadios.concat(sortKeyRadios, sortOrderRadios);
        for (var i = 0; i < arrangeRadios.length; i++) {
            arrangeRadios[i].onClick = function () {
                var sortKey = getSelectedKey(sortKeyRadios);
                var alongKey = getSelectedKey(alongRadios);
                var orderKey = getSelectedKey(sortOrderRadios);
                if (!sortKey || !alongKey || orderKey === null) return;
                applyArrangement(targetItems, sortKey, alongKey, orderKey);
                app.redraw();
            };
        }

        btnOk.onClick = function () {
            sortDialog.close(1);
        };
        btnCancel.onClick = function () {
            restoreItemPositions(savedPositions);
            app.redraw();
            sortDialog.close(0);
        };
        sortDialog.center();
        sortDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        var docSelection = app.activeDocument.selection;
        if (docSelection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var targetItems = [];
        for (var i = 0; i < docSelection.length; i++) {
            targetItems.push(docSelection[i]);
        }
        showSortDialog(targetItems, detectAlongDirection(docSelection), saveItemPositions(docSelection));
    }

    main();

})();
