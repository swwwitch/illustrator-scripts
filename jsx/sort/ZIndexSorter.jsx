#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

ZIndexSorter.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/sort/ZIndexSorter.jsx

### 概要：

- オブジェクトの重ね順を位置（X/Y）やZインデックスで並べ替える
- 昇順／降順／ランダムに並べ替え可能

### 主な機能：

- ダイアログで並べ替え軸と順序を指定
- キャンセル時には元の順序に復元
- 多言語対応（日本語／英語）

### 処理の流れ：

- 選択アイテムを収集
- 並べ替え軸・順序の設定をダイアログで取得
- 並べ替え後、Zインデックス順に再配置

### note：

- Illustrator 2025 以降で動作確認済

### 更新履歴：

- v1.0 (20250806) : 初期バージョン

*/

/*

### Script Name:

ZIndexSorter.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/sort/ZIndexSorter.jsx

### Overview:

- Sort object stacking order by position (X/Y) or Z-index
- Supports ascending / descending / random order

### Features:

- Dialog-based control of sorting axis and order
- Restore original order when canceled
- Bilingual support (Japanese / English)

### Workflow:

- Collect selected items
- Let user choose axis and order via dialog
- Apply Z-index reordering based on sorting

### note:

- Verified on Illustrator 2025 and later

### Change Log:

- v1.0 (20250806): Initial release
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.0";

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
var lang = getCurrentLang();

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
function reverseZOrder(sel) {
    if (!sel || sel.length < 2) {
        alert(LABELS.errors.selectMore[lang]);
        return;
    }
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        items.push(sel[i]);
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
function sortByAxis(sel, axis, order) {
    if (!sel || sel.length < 2) {
        alert(LABELS.errors.selectMore[lang]);
        return;
    }
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        items.push(sel[i]);
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
function sortByXAxis(sel, order) {
    sortByAxis(sel, "x", order);
}

/* Y軸で並べ替え / Sort by Y axis */
function sortByYAxis(sel, order) {
    sortByAxis(sel, "y", order);
}

/* 有効な選択を返す。なければ null / Return valid selection or null */
function getValidSelection() {
    if (app.documents.length === 0) {
        alert(LABELS.errors.noDocument[lang]);
        return null;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length < 2) {
        alert(LABELS.errors.selectMore[lang]);
        return null;
    }
    return sel;
}

function main() {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.alignChildren = "left";

    var doc = app.activeDocument;
    var sel = doc.selection;
    var originalOrder = [];
    for (var i = 0; i < sel.length; i++) {
        originalOrder.push(sel[i]);
    }

    var sortPanel = dlg.add("panel", undefined, LABELS.sortMethod[lang]);
    sortPanel.orientation = "column";
    sortPanel.alignChildren = "left";
    sortPanel.margins = [15, 20, 15, 10];
    var rbZOrder = sortPanel.add("radiobutton", undefined, LABELS.zOrder[lang]);
    var rbXAxis  = sortPanel.add("radiobutton", undefined, LABELS.xAxis[lang]);
    var rbYAxis  = sortPanel.add("radiobutton", undefined, LABELS.yAxis[lang]);
    rbZOrder.value = true;

    var orderPanel = dlg.add("panel", undefined, LABELS.orderMethod[lang]);
    orderPanel.orientation = "column";
    orderPanel.alignChildren = "left";
    orderPanel.margins = [15, 20, 15, 10];
    var rbAsc  = orderPanel.add("radiobutton", undefined, LABELS.asc[lang]);
    var rbDesc = orderPanel.add("radiobutton", undefined, LABELS.desc[lang]);
    var rbRand = orderPanel.add("radiobutton", undefined, LABELS.rand[lang]);
    rbAsc.value = true;

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = ["right", "bottom"];
    var btnCancel = btnGroup.add("button", undefined, LABELS.cancel[lang], {name: "cancel"});
    var btnOK = btnGroup.add("button", undefined, LABELS.ok[lang], {name: "ok"});

    var currentOrder = "asc";

    function applyPreviewWithOrder() {
        if (rbRand.value) {
            currentOrder = "rand";
        } else if (rbDesc.value) {
            currentOrder = "desc";
        } else {
            currentOrder = "asc";
        }

        var sel = getValidSelection();
        if (!sel) return;

        switch (true) {
            case rbZOrder.value:
                reverseZOrder(sel);
                break;
            case rbXAxis.value:
                sortByXAxis(sel, currentOrder);
                break;
            case rbYAxis.value:
                sortByYAxis(sel, currentOrder);
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
        dlg.close(1);
    };

    btnCancel.onClick = function() {
        if (originalOrder.length > 0) {
            reorderItems(originalOrder);
            app.redraw();
        }
        dlg.close(0);
    };

    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.97;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);

    dlg.show();
}

main();