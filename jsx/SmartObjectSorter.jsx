#target illustrator

/*
SmartObjectSorter.jsx
-------------------------------------------------------------
スクリプト概要：
選択したIllustratorオブジェクトを「高さ・幅・不透明度・カラー」で並び替えし、横または縦に揃えて整列・分布させるダイアログツールです。
整列基準・方向・順序・間隔の制御に加え、幅・高さを最大／最小に統一する機能を備えています。
プレビューを確認しながらリアルタイムに整列状態を調整できます。

Inspired by:
- m1b　https://community.adobe.com/t5/illustrator-discussions/script-that-sorts-items-in-selection/m-p/14413701#M396923
- John Wundes https://github.com/johnwun/js4ai/blob/master/organize.jsx

This software includes components developed by wundes.com and its contributors.
Copyright (c) 2005 wundes.com. All rights reserved.
Full license text available at: http://www.wundes.com/js4ai/copyright.txt

作成日：2024年6月3日
更新日：2025年6月4日
- 0.0.1: 初版リリース
- 0.0.2: 整列機能を調整
- 0.0.3: 縦方向中央揃えの不具合修正、初期プレビューの自動実行を無効化
- 0.0.4: 「幅／高さを揃える」機能を追加、UI整理
*/

// 並べ替え・整列ダイアログの適用処理で再帰的適用を防ぐフラグ
var skipApply = false;

// 並び方向の初期値（横並び）
var defaultAlong = "x";

// -------------------------------
// ラベル定義（日英対応）
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// Utility: check if there is a valid selection in the active document
function hasValidSelection() {
    return app.activeDocument &&
           app.activeDocument.selection &&
           app.activeDocument.selection.length > 0;
}

// Utility: get center X and Y of given bounds [left, top, right, bottom]
function getBoundsCenter(bounds) {
    var cx = (bounds[0] + bounds[2]) / 2;
    var cy = (bounds[1] + bounds[3]) / 2;
    return { x: cx, y: cy };
}

var lang = getCurrentLang();
var LABELS = {
    dirRandom: {
        ja: "ランダム",
        en: "Random"
    },
    matchMax: {
        ja: "最大",
        en: "Max"
    },
    matchMin: {
        ja: "最小",
        en: "Min"
    },
    dialogTitle: {
        ja: "オブジェクトの整列",
        en: "Object Alignment Tool"
    },
    alongX: {
        ja: "横並び",
        en: "Horizontal"
    },
    alongY: {
        ja: "縦並び",
        en: "Vertical"
    },
    byTitle: {
        ja: "基準",
        en: "Sort by"
    },
    byHeight: {
        ja: "高さ",
        en: "Height"
    },
    byWidth: {
        ja: "幅",
        en: "Width"
    },
    byOpacity: {
        ja: "不透明度",
        en: "Opacity"
    },
    byColor: {
        ja: "カラー",
        en: "Color"
    },
    byNumber: {
        ja: "数字",
        en: "Number"
    },
    directionTitle: {
        ja: "ソート順",
        en: "Sort Order"
    },
    dirAsc: {
        ja: "昇順",
        en: "Ascending"
    },
    dirDesc: {
        ja: "降順",
        en: "Descending"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnOk: {
        ja: "実行",
        en: "Apply"
    },
    alert_fail: {
        ja: "ファイルを開き、並べ替えるオブジェクトを選択してください。",
        en: "Open a file and select objects to sort."
    },
    alert_zorder_random: {
        ja: "ランダムと重ね順の組み合わせは安全でないため実行できません。",
        en: "Random and Z-order combination is unsafe and cannot be executed.",
    },
    verticalGroupTitle: {
        ja: "縦方向",
        en: "Vertical"
    },
    panelAlignVertical: {
        ja: "揃え",
        en: "Align Vertically"
    },
    alignVerticalLeft: {
        ja: "左",
        en: "Left"
    },
    alignVerticalCenter: {
        ja: "中央",
        en: "Center"
    },
    alignVerticalRight: {
        ja: "右",
        en: "Right"
    },
    panelSpacingVertical: {
        ja: "縦間隔",
        en: "Vertical Spacing"
    },
    // --- spacingVerticalNone inserted here ---
    spacingVerticalNone: {
        ja: "なし",
        en: "None"
    },
    spacingVerticalEven: {
        ja: "均等",
        en: "Even"
    },
    spacingVerticalZero: {
        ja: "ぴったり",
        en: "Tight"
    },
    spacingVerticalCustom: {
        ja: "指定",
        en: "Custom"
    },
    matchWidthTitle: {
        ja: "幅を揃える",
        en: "Match Width"
    },
    horizontalGroupTitle: {
        ja: "横方向",
        en: "Horizontal"
    },
    panelAlignHorizontal: {
        ja: "揃え",
        en: "Align Horizontally"
    },
    alignHorizontalTop: {
        ja: "上",
        en: "Top"
    },
    alignHorizontalMiddle: {
        ja: "中央",
        en: "Middle"
    },
    alignHorizontalBottom: {
        ja: "下",
        en: "Bottom"
    },
    panelSpacingHorizontal: {
        ja: "横間隔",
        en: "Horizontal Spacing"
    },
    spacingHorizontalEven: {
        ja: "均等",
        en: "Even"
    },
    spacingHorizontalZero: {
        ja: "ぴったり",
        en: "Tight"
    },
    spacingHorizontalCustom: {
        ja: "指定",
        en: "Custom"
    },
    matchHeightTitle: {
        ja: "高さを揃える",
        en: "Match Height"
    },
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview Bounds"
    }
};

// オブジェクト配列を指定方向に揃える
function alignObjects(selArr, alignType) {
    var i, j;
    switch (alignType) {
        case "top":
            // 上揃え：最も上の位置に合わせる
            var maxTop = -Infinity;
            for (i = 0; i < selArr.length; i++) {
                if (selArr[i].top > maxTop) {
                    maxTop = selArr[i].top;
                }
            }
            for (j = 0; j < selArr.length; j++) {
                selArr[j].top = maxTop;
            }
            break;
        case "bottom":
            // 下揃え：最も下の位置に合わせる
            var minBottom = Infinity;
            for (i = 0; i < selArr.length; i++) {
                var bottom = selArr[i].top - selArr[i].height;
                if (bottom < minBottom) {
                    minBottom = bottom;
                }
            }
            for (j = 0; j < selArr.length; j++) {
                selArr[j].top = minBottom + selArr[j].height;
            }
            break;
        case "left":
            // 左揃え：最も左の位置に合わせる
            var minLeft = Infinity;
            for (i = 0; i < selArr.length; i++) {
                if (selArr[i].left < minLeft) {
                    minLeft = selArr[i].left;
                }
            }
            for (j = 0; j < selArr.length; j++) {
                selArr[j].left = minLeft;
            }
            break;
        case "right":
            // 右揃え：最も右の位置に合わせる
            var maxRight = -Infinity;
            for (i = 0; i < selArr.length; i++) {
                var right = selArr[i].left + selArr[i].width;
                if (right > maxRight) {
                    maxRight = right;
                }
            }
            for (j = 0; j < selArr.length; j++) {
                selArr[j].left = maxRight - selArr[j].width;
            }
            break;
    }
}

var alert_fail, createItemSort, getInterval, numericSort, randomSort, rearrange, revNumericSort, attrParams, sortParams;

// 並び替え対象属性のマッピング（UIラジオボタンのhelpTip値と一致）
attrParams = {
    'h': "height", // 高さ
    'w': "width", // 幅
    'o': "opacity", // 不透明度
    'x': "left", // X座標
    'y': "top", // Y座標
    'color': "color" // カラー
};
numericSort = function(a, b) {
    return a - b;
};
revNumericSort = function(a, b) {
    return b - a;
};
randomSort = function(a, b) {
    return Math.random() - .5;
};
sortParams = {
    's': numericSort,
    "ns": numericSort,
    'l': revNumericSort,
    "ls": revNumericSort
};

// 配列を指定属性・方向で並べ替え、必要に応じて値を再割り当てする
rearrange = function(_arr, _byFilter, _sortDirection, _alongParameter, _interpolate) {
    if (_sortDirection === randomSort && _alongParameter === "zOrderPosition") {
        return; // Skip dangerous combination
    }
    var fin, max, selArrLen, start, tempAtts, _results;
    _arr.sort(_byFilter);
    selArrLen = _arr.length;
    tempAtts = [];
    while (selArrLen--) {
        tempAtts.push(_arr[selArrLen][_alongParameter]);
    }
    tempAtts.sort(_sortDirection);
    selArrLen = _arr.length;
    if (_alongParameter === "zOrderPosition") {
        if (_sortDirection === randomSort) {
            return; // Avoid risky random zOrder rearrangement
        }
        _results = [];
        while (selArrLen--) {
            _results.push((function() {
                var _results2 = [];
                var maxTries = 100;
                var tries = 0;
                while (_arr[selArrLen][_alongParameter] !== tempAtts[selArrLen] && tries < maxTries) {
                    if (_arr[selArrLen][_alongParameter] > tempAtts[selArrLen]) {
                        _arr[selArrLen].zOrder(ZOrderMethod.SENDBACKWARD);
                    } else {
                        _arr[selArrLen].zOrder(ZOrderMethod.BRINGFORWARD);
                    }
                    tries++;
                }
                return _results2;
            })());
        }
        return _results;
    } else {
        if (_alongParameter === 'top') {
            tempAtts.reverse();
        }
        start = tempAtts[0];
        fin = tempAtts[tempAtts.length - 1];
        max = _arr.length;
        for (var idx = 0; idx < max; idx++) {
            _arr[idx][_alongParameter] = (_sortDirection === randomSort || !_interpolate) ?
                tempAtts[idx] :
                getInterval(start, fin, max - 1, idx);
        }
    }
};

// 色のグレースケール値を取得
function getGrayScaleValue(breakdown) {
    if ('Array' !== breakdown.constructor.name)
        breakdown = getColorChannels(breakdown);

    var gray;

    if (breakdown.length === 4)
        gray = app.convertSampleColor(ImageColorSpace.CMYK, breakdown, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
    else if (breakdown.length === 3)
        gray = app.convertSampleColor(ImageColorSpace.RGB, breakdown, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
    else if (breakdown.length === 1)
        gray = breakdown;

    return gray;
}

function getColorChannels(col, tintFactor) {
    tintFactor = tintFactor || 1;

    if (col.hasOwnProperty('color'))
        col = col.color;

    if (col.constructor.name == 'SpotColor')
        col = col.spot.color;

    if (col.constructor.name === 'CMYKColor')
        return [col.cyan * tintFactor, col.magenta * tintFactor, col.yellow * tintFactor, col.black * tintFactor];
    else if (col.constructor.name === 'RGBColor')
        return [col.red * tintFactor, col.green * tintFactor, col.blue * tintFactor];
    else if (col.constructor.name === 'GrayColor')
        return [col.gray * tintFactor];

    return [0];
}

function getSortValue(obj, selectedAttr) {
    if (selectedAttr === "height") {
        return obj.height;
    } else if (selectedAttr === "width") {
        return obj.width;
    } else if (selectedAttr === "opacity") {
        return obj.opacity;
    } else if (selectedAttr === "left") {
        return obj.left;
    } else if (selectedAttr === "top") {
        return obj.top;
    } else if (selectedAttr === "color") {
        if (obj.fillColor) {
            return getGrayScaleValue(obj.fillColor);
        } else {
            return 0;
        }
    }
    return 0;
}

createItemSort = function(attrStr) {
    var attr = attrStr;
    return function(a, b) {
        // If sorting by color, use brightness
        if (attr === "color") {
            return getSortValue(a, "color") - getSortValue(b, "color");
        }
        return Number(a[attr]) - Number(b[attr]);
    };
};

// 均等配置時の間隔計算
getInterval = function(start, fin, len, curr) {
    var f, s;
    s = start;
    f = fin;
    if (start < 0) {
        s = 0;
        f = f - start;
    }
    return start + (((f - s) / len) * curr);
};
// ファイル未オープンまたは選択なし時のエラーメッセージ
alert_fail = LABELS.alert_fail[lang];

var dialog, byGroup, alongGroup, directionGroup;

// ラジオボタン配列から選択値（helpTip）を取得
function getSelectedValue(radioButtons) {
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].value) {
            return radioButtons[i].helpTip;
        }
    }
    return null;
}

// 並べ替え・整列の適用
function applyArrangement(selArr, byVal, alongVal, distVal, dirVal, alignVal) {
    if (skipApply) return;
    if (byVal === "n") {
        // 数字オプションでグループ内テキストから数値で並べ替え
        var sortableGroups = [];
        for (var i = 0; i < selArr.length; i++) {
            if (selArr[i].typename === "GroupItem") {
                var val = extractNumberFromGroup(selArr[i]);
                if (!isNaN(val)) {
                    sortableGroups.push({
                        group: selArr[i],
                        value: val
                    });
                }
            }
        }
        if (sortableGroups.length === 0) return;
        if (dirVal === "l") {
            sortableGroups.sort(function(a, b) { return b.value - a.value; });
        } else {
            sortableGroups.sort(function(a, b) { return a.value - b.value; });
        }
        var startTop = sortableGroups[0].group.top;
        var startLeft = sortableGroups[0].group.left;
        var spacing = 50;
        for (var j = 0; j < sortableGroups.length; j++) {
            sortableGroups[j].group.left = startLeft;
            sortableGroups[j].group.top = startTop - j * spacing;
        }
        return;
    }
    if (dirVal === 'r' && alongVal === 'z') {
        alert(LABELS.alert_zorder_random[lang]);
        return;
    }
    var sortFunc;
    if (dirVal === 's') {
        sortFunc = numericSort;
    } else if (dirVal === 'l') {
        sortFunc = revNumericSort;
    } else if (dirVal === 'r') {
        sortFunc = randomSort;
    } else {
        sortFunc = numericSort;
    }
    rearrange(selArr, createItemSort(attrParams[byVal]), sortFunc, attrParams[alongVal], distVal);
    // 並び方向に応じて整列処理
    if (alignVal === "none") {
        return; // 整列なし
    }
    if (alongVal === "y") {
        if (alignVal === "left") {
            alignObjects(selArr, "left");
        } else if (alignVal === "right") {
            alignObjects(selArr, "right");
        }
    } else {
        alignObjects(selArr, alignVal);
    }
}

// 並び方向自動判定用関数群
function calculateAutoTolerance(objects) {
    if (objects.length < 3) return 0;
    var bounds1 = objects[0].visibleBounds;
    var bounds2 = objects[1].visibleBounds;
    var bounds3 = objects[2].visibleBounds;
    var c1 = getBoundsCenter(bounds1);
    var c2 = getBoundsCenter(bounds2);
    var c3 = getBoundsCenter(bounds3);
    var dx1 = Math.abs(c1.x - c2.x);
    var dx2 = Math.abs(c2.x - c3.x);
    var dy1 = Math.abs(c1.y - c2.y);
    var dy2 = Math.abs(c2.y - c3.y);
    var avgDx = (dx1 + dx2) / 2;
    var avgDy = (dy1 + dy2) / 2;
    return Math.max(avgDx, avgDy) * 1.5;
}

function isTightlyAligned(values, tolerance) {
    var min = values[0];
    var max = values[0];
    for (var i = 1; i < values.length; i++) {
        if (values[i] < min) min = values[i];
        if (values[i] > max) max = values[i];
    }
    return (max - min) <= tolerance;
}

function detectAutoAlignment(objects, tolerance) {
    var xVals = [],
        yVals = [];
    for (var i = 0; i < objects.length; i++) {
        var bounds = objects[i].visibleBounds;
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        xVals.push(centerX);
        yVals.push(centerY);
    }
    var isVertical = isTightlyAligned(xVals, tolerance);
    var isHorizontal = isTightlyAligned(yVals, tolerance);
    if (isVertical && !isHorizontal) return "vertical";
    if (isHorizontal && !isVertical) return "horizontal";
    return "none";
}

// 並べ替え・整列ダイアログ作成

// メイン処理
function main() {
    var selArr = app.activeDocument.selection;
    // ダイアログ起動前に元の位置を保存
    var originalStates = [];
    var initialState = null;
    if (hasValidSelection()) {
        var selection = app.activeDocument.selection;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            originalStates.push({
                item: item,
                left: item.left,
                top: item.top
            });
        }
    }

    // 並び方向自動検出を実行（先頭3つの中心位置を使用）
    if (hasValidSelection() && app.activeDocument.selection.length >= 3) {
        var selection = app.activeDocument.selection;

        var bounds1 = selection[0].visibleBounds;
        var bounds2 = selection[1].visibleBounds;
        var bounds3 = selection[2].visibleBounds;
        var c1 = getBoundsCenter(bounds1);
        var c2 = getBoundsCenter(bounds2);
        var c3 = getBoundsCenter(bounds3);

        var dx1 = Math.abs(c1.x - c2.x);
        var dx2 = Math.abs(c2.x - c3.x);
        var dy1 = Math.abs(c1.y - c2.y);
        var dy2 = Math.abs(c2.y - c3.y);

        var avgDx = (dx1 + dx2) / 2;
        var avgDy = (dy1 + dy2) / 2;

        if (avgDx * 1.5 < avgDy) {
            defaultAlong = "y";
        } else {
            defaultAlong = "x";
        }
    }

    // 縦並びなら基準を「幅」、横並びなら「高さ」に初期化する
    var defaultBy = (defaultAlong === "y") ? "w" : "h";

    // キャンセル時に元の位置へ戻す処理をダイアログに追加
    var createDialog = function() {
        var result = false;
        if (!app.activeDocument) {
            alert(LABELS.alert_fail[lang]);
            return false;
        }
        var doc = activeDocument;
        var sel = doc.selection;
        if (sel.length === 0) {
            alert(LABELS.alert_fail[lang]);
            return false;
        }
        selArr = [];
        for (var i = 0; i < sel.length; i++) {
            selArr.push(sel[i]);
        }

        // Use single definition for all controls, matching the main dialog structure


        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);

        dialog.alignChildren = "left";
        dialog.orientation = "column";

        alongGroup = dialog.add("group");
        alongGroup.alignment = "center";
        alongGroup.orientation = "row";
        alongGroup.spacing = 5;
        alongGroup.margins = [10, 10, 10, 10];
        var leftAlignRadio = alongGroup.add("radiobutton", undefined, LABELS.alongX[lang]);
        var topAlignRadio = alongGroup.add("radiobutton", undefined, LABELS.alongY[lang]);
        if (defaultAlong === "y") {
            leftAlignRadio.value = false;
            topAlignRadio.value = true;
        } else {
            leftAlignRadio.value = true;
            topAlignRadio.value = false;
        }
        leftAlignRadio.helpTip = "x";
        topAlignRadio.helpTip = "y";
        var alongRadioButtons = [leftAlignRadio, topAlignRadio];
        var mainGroup = dialog.add("group");
        mainGroup.orientation = "row";
        mainGroup.alignChildren = ["fill", "top"];
        mainGroup.spacing = 10;
        var leftGroup = mainGroup.add("group");
        leftGroup.orientation = "column";
        leftGroup.alignChildren = "fill";
        leftGroup.spacing = 10;
        var rightGroup = mainGroup.add("group");
        rightGroup.orientation = "column";
        rightGroup.alignChildren = "fill";
        // rightGroup.spacing = 20;

        // 並べ替え基準パネル
        byGroup = leftGroup.add("panel", undefined, LABELS.byTitle[lang]);
        byGroup.orientation = "column";
        byGroup.alignChildren = "left";
        byGroup.spacing = 10;
        byGroup.margins = [15, 20, 15, 10];
        var heightRadio = byGroup.add("radiobutton", undefined, LABELS.byHeight[lang]);
        heightRadio.value = (defaultBy === "h");
        heightRadio.helpTip = "h";
        var widthRadio = byGroup.add("radiobutton", undefined, LABELS.byWidth[lang]);
        widthRadio.value = (defaultBy === "w");
        widthRadio.helpTip = "w";
        var opacityRadio = byGroup.add("radiobutton", undefined, LABELS.byOpacity[lang]);
        opacityRadio.value = false;
        opacityRadio.helpTip = "o";
        var colorRadio = byGroup.add("radiobutton", undefined, LABELS.byColor[lang]);
        colorRadio.value = false;
        colorRadio.helpTip = "color";
        var numberRadio = byGroup.add("radiobutton", undefined, LABELS.byNumber[lang]);
        numberRadio.value = false;
        numberRadio.helpTip = "n";
        var byRadioButtons = [heightRadio, widthRadio, opacityRadio, colorRadio, numberRadio];

        // 並び順パネル
        directionGroup = leftGroup.add("panel", undefined, LABELS.directionTitle[lang]);
        directionGroup.orientation = "column";
        directionGroup.alignChildren = "left";
        directionGroup.spacing = 5;
        directionGroup.margins = [15, 20, 15, 10];
        var ascendingRadio = directionGroup.add("radiobutton", undefined, LABELS.dirAsc[lang]);
        ascendingRadio.value = true;
        ascendingRadio.helpTip = "s";
        var descendingRadio = directionGroup.add("radiobutton", undefined, LABELS.dirDesc[lang]);
        descendingRadio.value = false;
        descendingRadio.helpTip = "l";
        var randomRadio = directionGroup.add("radiobutton", undefined, LABELS.dirRandom[lang]);
        randomRadio.value = false;
        randomRadio.helpTip = "r";
        var directionRadioButtons = [ascendingRadio, descendingRadio, randomRadio];



        // 中央ペイン(panel) - 縦方向
        var verticalAlignPanelContainer = mainGroup.add("panel", undefined, "");
        verticalAlignPanelContainer.alignChildren = "fill";
        verticalAlignPanelContainer.margins = [10, 20, 10, 10];
        verticalAlignPanelContainer.text = LABELS.verticalGroupTitle[lang];

        // 縦方向：揃えパネル
        var verticalAlignPanel = verticalAlignPanelContainer.add("panel", undefined, LABELS.panelAlignVertical[lang]);
        verticalAlignPanel.orientation = "row";
        verticalAlignPanel.alignChildren = "center";
        verticalAlignPanel.margins = [10, 20, 10, 10];

        var alignLeftBtnVerticalSmart = verticalAlignPanel.add("radiobutton", undefined, LABELS.alignVerticalLeft[lang]);
        var alignCenterBtnVerticalSmart = verticalAlignPanel.add("radiobutton", undefined, LABELS.alignVerticalCenter[lang]);
        var alignRightBtnVerticalSmart = verticalAlignPanel.add("radiobutton", undefined, LABELS.alignVerticalRight[lang]);

        alignLeftBtnVerticalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "left", false);
            app.redraw();
        };
        alignCenterBtnVerticalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "center", false);
            app.redraw();
        };
        alignRightBtnVerticalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "right", false);
            app.redraw();
        };

        // 縦方向: 縦間隔パネル
        createSpacingControlGroup(verticalAlignPanelContainer, {
            panelTitle: LABELS.panelSpacingVertical[lang],
            none: LABELS.spacingVerticalNone[lang],
            even: LABELS.spacingVerticalEven[lang],
            zero: LABELS.spacingVerticalZero[lang],
            custom: LABELS.spacingVerticalCustom[lang]
        }, selArr, false);

        // 幅を揃えるパネル
        var matchWidthPanel = verticalAlignPanelContainer.add("panel", undefined, LABELS.matchWidthTitle[lang]);
        matchWidthPanel.orientation = "row";
        matchWidthPanel.alignChildren = "center";
        matchWidthPanel.margins = [10, 20, 10, 10];

        var matchMaxWidthRadio = matchWidthPanel.add("radiobutton", undefined, LABELS.matchMax[lang]);
        var matchMinWidthRadio = matchWidthPanel.add("radiobutton", undefined, LABELS.matchMin[lang]);
        matchMaxWidthRadio.value = false;
        matchMinWidthRadio.value = false;

        // 最大幅で統一
        matchMaxWidthRadio.onClick = function() {
            var maxWidth = 0;
            for (var i = 0; i < selArr.length; i++) {
                if (selArr[i].width > maxWidth) {
                    maxWidth = selArr[i].width;
                }
            }
            for (var i = 0; i < selArr.length; i++) {
                var ratio = maxWidth / selArr[i].width;
                selArr[i].resize(ratio * 100, 100);
            }
            app.redraw();
        };
        // 最小幅で統一
        matchMinWidthRadio.onClick = function() {
            var minWidth = Infinity;
            for (var i = 0; i < selArr.length; i++) {
                if (selArr[i].width < minWidth) {
                    minWidth = selArr[i].width;
                }
            }
            for (var i = 0; i < selArr.length; i++) {
                var ratio = minWidth / selArr[i].width;
                selArr[i].resize(ratio * 100, 100);
            }
            app.redraw();
        };



        // 右ペイン(panel) - 横方向
        var horizontalAlignPanelContainer = mainGroup.add("panel", undefined, "");
        horizontalAlignPanelContainer.alignChildren = "fill";
        horizontalAlignPanelContainer.margins = [10, 20, 10, 10];
        horizontalAlignPanelContainer.text = LABELS.horizontalGroupTitle[lang];

        // 横方向：揃えパネル
        var horizontalAlignPanel = horizontalAlignPanelContainer.add("panel", undefined, LABELS.panelAlignHorizontal[lang]);
        horizontalAlignPanel.orientation = "row";
        horizontalAlignPanel.alignChildren = "center";
        horizontalAlignPanel.margins = [10, 20, 10, 10];

        var alignTopBtnHorizontalSmart = horizontalAlignPanel.add("radiobutton", undefined, LABELS.alignHorizontalTop[lang]);
        var alignMiddleBtnHorizontalSmart = horizontalAlignPanel.add("radiobutton", undefined, LABELS.alignHorizontalMiddle[lang]);
        var alignBottomBtnHorizontalSmart = horizontalAlignPanel.add("radiobutton", undefined, LABELS.alignHorizontalBottom[lang]);

        alignTopBtnHorizontalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "top", true);
            app.redraw();
        };
        alignMiddleBtnHorizontalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "middle", true);
            app.redraw();
        };
        alignBottomBtnHorizontalSmart.onClick = function() {
            smartPreviewAlignGeneric(selArr, "bottom", true);
            app.redraw();
        };

        // 横間隔パネル
        var horizontalSpacingPanel = horizontalAlignPanelContainer.add("panel", undefined, LABELS.panelSpacingHorizontal[lang]);
        horizontalSpacingPanel.orientation = "column";
        horizontalSpacingPanel.alignChildren = "left";
        horizontalSpacingPanel.margins = [10, 20, 10, 10];

        var spacingHorizontalRadioGroup = horizontalSpacingPanel.add("group");
        spacingHorizontalRadioGroup.orientation = "column";
        spacingHorizontalRadioGroup.alignChildren = "left";

        // Remove "None" option for horizontal spacing
        var spacingEvenBtnHorizontalSmart = spacingHorizontalRadioGroup.add("radiobutton", undefined, LABELS.spacingHorizontalEven[lang]);
        var spacingZeroBtnHorizontalSmart = spacingHorizontalRadioGroup.add("radiobutton", undefined, LABELS.spacingHorizontalZero[lang]);

        var spacingHorizontalCustomGroup = spacingHorizontalRadioGroup.add("group");
        spacingHorizontalCustomGroup.orientation = "row";
        spacingHorizontalCustomGroup.alignChildren = "left";
        var spacingCustomBtnHorizontalSmart = spacingHorizontalCustomGroup.add("radiobutton", undefined, LABELS.spacingHorizontalCustom[lang]);
        var spacingInputHorizontalSmart = spacingHorizontalCustomGroup.add("edittext", undefined, "20");
        spacingInputHorizontalSmart.characters = 5;
        spacingInputHorizontalSmart.enabled = false;

        // Set "even" as default selection
        // spacingEvenBtnHorizontalSmart.value = true;
        spacingEvenBtnHorizontalSmart.value = false;
        spacingZeroBtnHorizontalSmart.value = false;
        spacingCustomBtnHorizontalSmart.value = false;

        spacingEvenBtnHorizontalSmart.onClick = function() {
            onSmartSpacingRadioClick(
                null,
                spacingEvenBtnHorizontalSmart,
                spacingZeroBtnHorizontalSmart,
                spacingCustomBtnHorizontalSmart,
                spacingInputHorizontalSmart,
                "even",
                selArr,
                true
            );
        };
        spacingZeroBtnHorizontalSmart.onClick = function() {
            onSmartSpacingRadioClick(
                null,
                spacingEvenBtnHorizontalSmart,
                spacingZeroBtnHorizontalSmart,
                spacingCustomBtnHorizontalSmart,
                spacingInputHorizontalSmart,
                "zero",
                selArr,
                true
            );
        };
        spacingCustomBtnHorizontalSmart.onClick = function() {
            onSmartSpacingRadioClick(
                null,
                spacingEvenBtnHorizontalSmart,
                spacingZeroBtnHorizontalSmart,
                spacingCustomBtnHorizontalSmart,
                spacingInputHorizontalSmart,
                "custom",
                selArr,
                true
            );
        };
        spacingInputHorizontalSmart.onChange = function() {
            if (spacingCustomBtnHorizontalSmart.value) {
                smartPreviewSpacingGeneric(selArr, spacingInputHorizontalSmart, "custom", true);
                app.redraw();
            }
        };

        // 高さを揃えるパネル（ダイアログ下部、キャンセル／OKの直上）
        var matchHeightPanel = horizontalAlignPanelContainer.add("panel", undefined, LABELS.matchHeightTitle[lang]);
        matchHeightPanel.orientation = "row";
        matchHeightPanel.alignChildren = "center";
        matchHeightPanel.margins = [10, 20, 10, 10];

        var matchMaxHeightRadio = matchHeightPanel.add("radiobutton", undefined, LABELS.matchMax[lang]);
        var matchMinHeightRadio = matchHeightPanel.add("radiobutton", undefined, LABELS.matchMin[lang]);
        matchMaxHeightRadio.value = false;
        matchMinHeightRadio.value = false;

        // 最大高さで統一（幅はそのまま）
        matchMaxHeightRadio.onClick = function() {
            var maxHeight = 0;
            for (var i = 0; i < selArr.length; i++) {
                if (selArr[i].height > maxHeight) {
                    maxHeight = selArr[i].height;
                }
            }
            for (var i = 0; i < selArr.length; i++) {
                var ratio = maxHeight / selArr[i].height;
                selArr[i].resize(100, ratio * 100);
            }
            app.redraw();
        };
        // 最小高さで統一（幅はそのまま）
        matchMinHeightRadio.onClick = function() {
            var minHeight = Infinity;
            for (var i = 0; i < selArr.length; i++) {
                if (selArr[i].height < minHeight) {
                    minHeight = selArr[i].height;
                }
            }
            for (var i = 0; i < selArr.length; i++) {
                var ratio = minHeight / selArr[i].height;
                selArr[i].resize(100, ratio * 100);
            }
            app.redraw();
        };


        // 下部行: スペーサ、キャンセル・OKボタン（previewBoundsCheckboxは中央ペイン下部に移動）
        var bottomRowGroup = dialog.add("group");
        bottomRowGroup.orientation = "row";
        bottomRowGroup.alignment = ["fill", "bottom"];
        bottomRowGroup.alignChildren = ["left", "center"];
        bottomRowGroup.margins = [10, 10, 10, 10];


        // プレビュー境界チェックボックスを中央ペイン下部に追加
        var previewBoundsCheckbox = bottomRowGroup.add("checkbox", undefined, LABELS.previewBounds[lang]);
        previewBoundsCheckbox.value = false;

        // Flexible spacer
        var spacer = bottomRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 100;
        spacer.maximumSize.height = 0;

        // Cancel button
        var cancelBtn = bottomRowGroup.add("button", undefined, LABELS.btnCancel[lang], {
            name: "cancel"
        });

        // OK button
        var okBtn = bottomRowGroup.add("button", undefined, LABELS.btnOk[lang], {
            name: "ok"
        });
        okBtn.active = true;

        function updatePreview(forceRedraw) {
            var byVal = getSelectedValue(byRadioButtons);
            var alongVal = getSelectedValue(alongRadioButtons);
            var distVal = false;
            var dirVal = getSelectedValue(directionRadioButtons);
            // No alignment panel, so alignVal is always "none"
            var alignVal = "none";
            if (!byVal || !alongVal || dirVal === null) {
                return;
            }
            applyArrangement(selArr, byVal, alongVal, distVal, dirVal, alignVal);
            if (forceRedraw !== false) {
                app.redraw();
            }
        }

        function addListeners(radioButtons) {
            for (var i = 0; i < radioButtons.length; i++) {
                radioButtons[i].onClick = function() {
                    updatePreview(true);
                };
            }
        }
        addListeners(byRadioButtons);
        addListeners(directionRadioButtons);

        for (var i = 0; i < alongRadioButtons.length; i++) {
            alongRadioButtons[i].onClick = function() {
                updatePreview(true);
            };
        }
        // UIパーツの定義完了後に並び方向の初期状態を反映
        // updatePreview(true);

        okBtn.onClick = function() {
            dialog.close(1);
        };
        cancelBtn.onClick = function() {
            for (var i = 0; i < originalStates.length; i++) {
                var obj = originalStates[i];
                obj.item.left = obj.left;
                obj.item.top = obj.top;
            }
            app.redraw();
            dialog.close(0);
        };
        dialog.center();
        var resultShow = dialog.show();
        if (resultShow === 1) {
            // すでにプレビューで反映済みなので何もせずtrueを返す
            return true;
        } else {
            return false;
        }
    };

    createDialog();
}


// 間隔プレビュー関数
function smartPreviewSpacingGeneric(selArr, spacingInput, spacingType, isHorizontal) {
    var spacingValue = 0;
    if (spacingType === "custom" && spacingInput && spacingInput.text !== "") {
        spacingValue = parseFloat(spacingInput.text);
    }
    var sel = selArr.slice();
    sel.sort(function(a, b) {
        return isHorizontal ? a.left - b.left : b.top - a.top;
    });

    // Vertical "even" spacing logic
    if (!isHorizontal && spacingType === "even") {
        var sorted = sel.slice();
        sorted.sort(function(a, b) {
            return b.top - a.top;
        });

        var topMost = sorted[0].top;
        var bottomMost = sorted[sorted.length - 1].top - sorted[sorted.length - 1].height;

        var totalHeight = 0;
        for (var i = 0; i < sorted.length; i++) {
            totalHeight += sorted[i].height;
        }

        var totalGap = topMost - bottomMost - totalHeight;
        var gap = totalGap / (sorted.length - 1);

        var y = topMost;
        for (var j = 0; j < sorted.length; j++) {
            sorted[j].top = y;
            y -= sorted[j].height + gap;
        }
        return;
    }

    // Horizontal "even" spacing logic
    if (isHorizontal && spacingType === "even") {
        var sorted = sel.slice();
        sorted.sort(function(a, b) {
            return a.left - b.left;
        });

        var leftMost = sorted[0].left;
        var rightMost = sorted[sorted.length - 1].left + sorted[sorted.length - 1].width;

        var totalWidth = 0;
        for (var i = 0; i < sorted.length; i++) {
            totalWidth += sorted[i].width;
        }

        var totalGap = rightMost - leftMost - totalWidth;
        var gap = totalGap / (sorted.length - 1);

        var x = leftMost;
        for (var j = 0; j < sorted.length; j++) {
            sorted[j].left = x;
            x += sorted[j].width + gap;
        }
        return;
    }

    if (!isHorizontal && spacingType === "custom") {
        var sorted = sel.slice();
        sorted.sort(function(a, b) {
            return b.top - a.top;
        });

        var y = sorted[0].top;
        for (var i = 0; i < sorted.length; i++) {
            sorted[i].top = y;
            y -= sorted[i].height + spacingValue;
        }
        return;
    }

    var currentPos = isHorizontal ? sel[0].left : sel[0].top;
    for (var i = 0; i < sel.length; i++) {
        if (i !== 0) {
            currentPos += spacingValue;
        }
        if (isHorizontal) {
            sel[i].left = currentPos;
            currentPos += sel[i].width;
        } else {
            sel[i].top = currentPos;
            currentPos -= sel[i].height;
        }
    }
}

// ラジオボタンクリック時の処理関数
function onSmartSpacingRadioClick(noneBtn, evenBtn, zeroBtn, customBtn, customInput, type, selArr, isHorizontal) {
    if (noneBtn) {
        noneBtn.value = (type === "none");
    }
    evenBtn.value = (type === "even");
    zeroBtn.value = (type === "zero");
    customBtn.value = (type === "custom");
    customInput.enabled = (type === "custom");

    if (type === "even" || type === "zero" || type === "custom") {
        smartPreviewSpacingGeneric(selArr, customInput, type, isHorizontal);
        app.redraw();
    }
}

// 間隔制御グループを生成する関数（縦・横共通）
function createSpacingControlGroup(parentPanel, labels, selArr, isHorizontal) {
    var spacingPanel = parentPanel.add("panel", undefined, labels.panelTitle);
    spacingPanel.orientation = "column";
    spacingPanel.alignChildren = "left";
    spacingPanel.margins = [10, 20, 10, 10];

    var spacingRadioGroup = spacingPanel.add("group");
    spacingRadioGroup.orientation = "column";
    spacingRadioGroup.alignChildren = "left";

    // Removed: var spacingNoneBtn = spacingRadioGroup.add("radiobutton", undefined, labels.none);
    var spacingEvenBtn = spacingRadioGroup.add("radiobutton", undefined, labels.even);
    var spacingZeroBtn = spacingRadioGroup.add("radiobutton", undefined, labels.zero);

    var spacingCustomGroup = spacingRadioGroup.add("group");
    spacingCustomGroup.orientation = "row";
    spacingCustomGroup.alignChildren = "left";

    var spacingCustomBtn = spacingCustomGroup.add("radiobutton", undefined, labels.custom);
    var spacingInput = spacingCustomGroup.add("edittext", undefined, "20");
    spacingInput.characters = 5;
    spacingInput.enabled = false;

    // Removed: spacingNoneBtn.value = true;

    // Removed: spacingNoneBtn.onClick handler

    spacingEvenBtn.onClick = function() {
        onSmartSpacingRadioClick(
            null,
            spacingEvenBtn,
            spacingZeroBtn,
            spacingCustomBtn,
            spacingInput,
            "even",
            selArr,
            isHorizontal
        );
    };
    spacingZeroBtn.onClick = function() {
        onSmartSpacingRadioClick(
            null,
            spacingEvenBtn,
            spacingZeroBtn,
            spacingCustomBtn,
            spacingInput,
            "zero",
            selArr,
            isHorizontal
        );
    };
    spacingCustomBtn.onClick = function() {
        onSmartSpacingRadioClick(
            null,
            spacingEvenBtn,
            spacingZeroBtn,
            spacingCustomBtn,
            spacingInput,
            "custom",
            selArr,
            isHorizontal
        );
    };
    spacingInput.onChange = function() {
        if (spacingCustomBtn.value) {
            smartPreviewSpacingGeneric(selArr, spacingInput, "custom", isHorizontal);
            app.redraw();
        }
    };
}

// 整列方向（横または縦）のプレビュー整列処理関数
function smartPreviewAlignGeneric(sel, alignType, isHorizontal) {
    if (!sel || sel.length === 0) return;

    // --- Patch: previewBoundsCheckbox affects positioning ---
    var usePreviewBounds = (typeof previewBoundsCheckbox !== "undefined" && previewBoundsCheckbox.value);
    var objBounds = [];
    for (var i = 0; i < sel.length; i++) {
        objBounds[i] = usePreviewBounds ? sel[i].visibleBounds : sel[i].geometricBounds;
    }

    // Compute alignment positions
    var positions = [];
    for (var i = 0; i < sel.length; i++) {
        var bounds = objBounds[i];
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        if (isHorizontal) {
            if (alignType === "top") {
                positions.push(bounds[1]);
            } else if (alignType === "middle") {
                positions.push(bounds[1] - (height / 2));
            } else if (alignType === "bottom") {
                positions.push(bounds[1] - height);
            }
        } else {
            if (alignType === "left") {
                positions.push(bounds[0]);
            } else if (alignType === "center") {
                positions.push(bounds[0] + (width / 2));
            } else if (alignType === "right") {
                positions.push(bounds[0] + width);
            }
        }
    }

    var targetPos;
    if (positions.length === 0) return;
    if (alignType === "top" || alignType === "left") {
        targetPos = Math.min.apply(null, positions);
    } else if (alignType === "bottom" || alignType === "right") {
        targetPos = Math.max.apply(null, positions);
    } else if (alignType === "middle" || alignType === "center") {
        var sum = 0;
        for (var i = 0; i < positions.length; i++) {
            sum += positions[i];
        }
        targetPos = sum / positions.length;
    }

    // --- Patch: apply offsets so previewBoundsCheckbox affects positioning ---
    for (var k = 0; k < sel.length; k++) {
        var bounds = objBounds[k];
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        var dx = sel[k].left - bounds[0];
        var dy = sel[k].top - bounds[1];

        if (isHorizontal) {
            if (alignType === "top") {
                sel[k].top = targetPos + dy;
            } else if (alignType === "middle") {
                sel[k].top = targetPos + height / 2 + dy;
            } else if (alignType === "bottom") {
                sel[k].top = targetPos + height + dy;
            }
        } else {
            if (alignType === "left") {
                sel[k].left = targetPos + dx;
            } else if (alignType === "center") {
                sel[k].left = targetPos - width / 2 + dx;
            } else if (alignType === "right") {
                sel[k].left = targetPos - width + dx;
            }
        }
    }
}

main();
// グループ内の最初の数字テキストを抽出する関数
function extractNumberFromGroup(group) {
    var items = group.pageItems;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "TextFrame") {
            var text = item.contents;
            if (/^\d+$/.test(text)) {
                return parseFloat(text);
            }
        } else if (item.typename === "GroupItem") {
            var nestedVal = extractNumberFromGroup(item);
            if (!isNaN(nestedVal)) return nestedVal;
        }
    }
    return NaN;
}