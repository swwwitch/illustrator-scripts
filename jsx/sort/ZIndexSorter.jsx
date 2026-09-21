#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの重ね順を、指定した基準で並べ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md

### Overview

Reorders the stacking order of the selected objects according to a chosen criterion.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ZIndexSorter";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var LABELS = {
        dialogTitle: {
            ja: "重ね順ソート " + SCRIPT_VERSION,
            en: "Z-Index Sorter " + SCRIPT_VERSION
        },
        sortMethod: {
            ja: "ソート方法",
            en: "Sort Method"
        },
        orderMethod: {
            ja: "並び順",
            en: "Order"
        },
        zOrder: {
            ja: "現在の重ね順",
            en: "Current Z-Order"
        },
        xAxis: {
            ja: "X軸",
            en: "X Axis"
        },
        yAxis: {
            ja: "Y軸",
            en: "Y Axis"
        },
        asc: {
            ja: "昇順",
            en: "Ascending"
        },
        desc: {
            ja: "降順",
            en: "Descending"
        },
        rand: {
            ja: "ランダム",
            en: "Random"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        tipZOrder: { ja: "いまの重ね順をそのまま基準にします。並び順だけを変えたいときに使います。", en: "Uses the current stacking order as the basis. Pick this when only the direction should change." },
        tipXAxis:  { ja: "X座標を基準に重ね順を組み直します。", en: "Restacks the objects by their X position." },
        tipYAxis:  { ja: "Y座標を基準に重ね順を組み直します。", en: "Restacks the objects by their Y position." },
        tipAsc:    { ja: "基準の値が小さいものほど背面にします。", en: "Puts objects with smaller values further back." },
        tipDesc:   { ja: "基準の値が大きいものほど背面にします。", en: "Puts objects with larger values further back." },
        tipRand:   { ja: "基準と関係なく、重ね順をシャッフルします。", en: "Shuffles the stacking order regardless of the basis." },
        errors: {
            selectMore: {
                ja: "2つ以上のオブジェクトを選択してください。",
                en: "Please select two or more objects."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            }
        }
    };

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* アイテムを基準アイテムの前に順に配置 / Reorder items based on the first element */
    function reorderItems(items) {
        var baseItem = items[0];
        for (var j = items.length - 1; j >= 0; j--) {
            if (items[j] !== baseItem) {
                items[j].move(baseItem, ElementPlacement.PLACEBEFORE);
            }
        }
    }

    /* 選択アイテムの重ね順を逆転 / Reverse Z-order of selected items */
    function reverseZOrder(currentSelection) {
        if (!currentSelection || currentSelection.length < 2) {
            alert(LABELS.errors.selectMore[uiLang]);
            return;
        }
        var items = [];
        for (var i = 0; i < currentSelection.length; i++) {
            items.push(currentSelection[i]);
        }
        reorderItems(items);
    }

    /* Fisher-Yatesアルゴリズムで配列をシャッフル / Shuffle array with Fisher-Yates */
    function fisherYatesShuffle(array) {
        for (var i = array.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = array[i];
            array[i] = array[j];
            array[j] = temp;
        }
        return array;
    }

    /* 指定軸で選択アイテムを並べ替え / Sort selected items by specified axis */
    function sortByAxis(currentSelection, axis, order) {
        if (!currentSelection || currentSelection.length < 2) {
            alert(LABELS.errors.selectMore[uiLang]);
            return;
        }
        var items = [];
        for (var i = 0; i < currentSelection.length; i++) {
            items.push(currentSelection[i]);
        }

        if (order === "rand") {
            items = fisherYatesShuffle(items);
        } else {
            items.sort(function(a, b) {
                var aVal = (axis === "x") ? a.geometricBounds[0] : a.geometricBounds[1];
                var bVal = (axis === "x") ? b.geometricBounds[0] : b.geometricBounds[1];
                if (axis === "y") {
                    return (order === "desc") ? aVal - bVal : bVal - aVal;
                } else {
                    return (order === "desc") ? bVal - aVal : aVal - bVal;
                }
            });
        }

        reorderItems(items);
    }

    /* X軸で並べ替え / Sort by X axis */
    function sortByXAxis(currentSelection, order) {
        sortByAxis(currentSelection, "x", order);
    }

    /* Y軸で並べ替え / Sort by Y axis */
    function sortByYAxis(currentSelection, order) {
        sortByAxis(currentSelection, "y", order);
    }

    /* 有効な選択を返す。なければ null / Return valid selection or null */
    function getValidSelection() {
        if (app.documents.length === 0) {
            alert(LABELS.errors.noDocument[uiLang]);
            return null;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length < 2) {
            alert(LABELS.errors.selectMore[uiLang]);
            return null;
        }
        return currentSelection;
    }

    function main() {
        var dialog = new Window("dialog", LABELS.dialogTitle[uiLang]);
        dialog.alignChildren = "left";

        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        var originalOrder = [];
        for (var i = 0; i < currentSelection.length; i++) {
            originalOrder.push(currentSelection[i]);
        }

        var sortPanel = dialog.add("panel", undefined, LABELS.sortMethod[uiLang]);
        sortPanel.orientation = "column";
        sortPanel.alignChildren = "left";
        sortPanel.margins = [15, 20, 15, 10];
        var rbZOrder = sortPanel.add("radiobutton", undefined, LABELS.zOrder[uiLang]);
        rbZOrder.helpTip = LABELS.tipZOrder[uiLang];
        var rbXAxis  = sortPanel.add("radiobutton", undefined, LABELS.xAxis[uiLang]);
        rbXAxis.helpTip = LABELS.tipXAxis[uiLang];
        var rbYAxis  = sortPanel.add("radiobutton", undefined, LABELS.yAxis[uiLang]);
        rbYAxis.helpTip = LABELS.tipYAxis[uiLang];
        rbZOrder.value = true;

        var orderPanel = dialog.add("panel", undefined, LABELS.orderMethod[uiLang]);
        orderPanel.orientation = "column";
        orderPanel.alignChildren = "left";
        orderPanel.margins = [15, 20, 15, 10];
        var rbAsc  = orderPanel.add("radiobutton", undefined, LABELS.asc[uiLang]);
        rbAsc.helpTip = LABELS.tipAsc[uiLang];
        var rbDesc = orderPanel.add("radiobutton", undefined, LABELS.desc[uiLang]);
        rbDesc.helpTip = LABELS.tipDesc[uiLang];
        var rbRand = orderPanel.add("radiobutton", undefined, LABELS.rand[uiLang]);
        rbRand.helpTip = LABELS.tipRand[uiLang];
        rbAsc.value = true;

        var btnGroup = dialog.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = ["right", "bottom"];
        var btnCancel = btnGroup.add("button", undefined, LABELS.cancel[uiLang], {name: "cancel"});
        var btnOK = btnGroup.add("button", undefined, LABELS.ok[uiLang], {name: "ok"});

        var currentOrder = "asc";

        function applyPreviewWithOrder() {
            if (rbRand.value) {
                currentOrder = "rand";
            } else if (rbDesc.value) {
                currentOrder = "desc";
            } else {
                currentOrder = "asc";
            }

            var currentSelection = getValidSelection();
            if (!currentSelection) return;

            switch (true) {
                case rbZOrder.value:
                    reverseZOrder(currentSelection);
                    break;
                case rbXAxis.value:
                    sortByXAxis(currentSelection, currentOrder);
                    break;
                case rbYAxis.value:
                    sortByYAxis(currentSelection, currentOrder);
                    break;
            }
            app.redraw();
        }

        rbZOrder.onClick = applyPreviewWithOrder;
        rbXAxis.onClick  = applyPreviewWithOrder;
        rbYAxis.onClick  = applyPreviewWithOrder;
        rbAsc.onClick    = applyPreviewWithOrder;
        rbDesc.onClick   = applyPreviewWithOrder;
        rbRand.onClick   = applyPreviewWithOrder;

        btnOK.onClick = function() {
            dialog.close(1);
        };

        btnCancel.onClick = function() {
            if (originalOrder.length > 0) {
                reorderItems(originalOrder);
                app.redraw();
            }
            dialog.close(0);
        };

        var offsetX = 300;
        var offsetY = 0;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dialog, offsetX, offsetY) {
            dialog.onShow = function () {
                var currentX = dialog.location[0];
                var currentY = dialog.location[1];
                dialog.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dialog, opacityValue) {
            dialog.opacity = opacityValue;
        }

        setDialogOpacity(dialog, dialogOpacity);
        shiftDialogPosition(dialog, offsetX, offsetY);

        dialog.show();
    }

    main();

})();
