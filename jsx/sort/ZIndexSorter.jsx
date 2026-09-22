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
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X = 300;  /* 表示位置を右へずらす量 / shift the dialog right by this much */
    var DIALOG_OFFSET_Y = 0;    /* 表示位置を下へずらす量 / shift the dialog down by this much */
    var DIALOG_OPACITY = 0.97;  /* ダイアログの不透明度 / dialog opacity */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "重ね順ソート", en: "Z-Index Sorter" }
        },
        panel: {
            sortMethod: { ja: "ソート方法", en: "Sort Method" },
            orderMethod: { ja: "並び順", en: "Order" }
        },
        radio: {
            zOrder: { ja: "現在の重ね順", en: "Current Z-Order" },
            xAxis: { ja: "X軸", en: "X Axis" },
            yAxis: { ja: "Y軸", en: "Y Axis" },
            asc: { ja: "昇順", en: "Ascending" },
            desc: { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" }
        },
        tooltip: {
            zOrder: {
                ja: "いまの重ね順をそのまま基準にします。並び順だけを変えたいときに使います。",
                en: "Uses the current stacking order as the basis. Pick this when only the direction should change."
            },
            xAxis: { ja: "X座標を基準に重ね順を組み直します。", en: "Restacks the objects by their X position." },
            yAxis: { ja: "Y座標を基準に重ね順を組み直します。", en: "Restacks the objects by their Y position." },
            asc: { ja: "基準の値が小さいものほど背面にします。", en: "Puts objects with smaller values further back." },
            desc: { ja: "基準の値が大きいものほど背面にします。", en: "Puts objects with larger values further back." },
            random: { ja: "基準と関係なく、重ね順をシャッフルします。", en: "Shuffles the stacking order regardless of the basis." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectMore: { ja: "2つ以上のオブジェクトを選択してください。", en: "Please select two or more objects." },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
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
    // 重ね順 / Stacking order
    // =========================================

    /**
     * 先頭のオブジェクトの前面へ、末尾から順に移す（配列の先頭ほど背面になる）
     * @param {PageItem[]} orderedItems - 並べたい順のオブジェクト
     * @returns {void}
     */
    function reorderItems(orderedItems) {
        var baseItem = orderedItems[0];
        for (var j = orderedItems.length - 1; j >= 0; j--) {
            if (orderedItems[j] !== baseItem) {
                orderedItems[j].move(baseItem, ElementPlacement.PLACEBEFORE);
            }
        }
    }

    /**
     * 選択を JavaScript の配列に写す
     * @param {PageItem[]} docSelection - 選択オブジェクト
     * @returns {PageItem[]} 写した配列
     */
    function copyItems(docSelection) {
        var copiedItems = [];
        for (var i = 0; i < docSelection.length; i++) {
            copiedItems.push(docSelection[i]);
        }
        return copiedItems;
    }

    /**
     * 配列の順序をシャッフルする（Fisher–Yates、直接書き換える）
     * @param {PageItem[]} targetItems - 対象の配列
     * @returns {void}
     */
    function shuffleItems(targetItems) {
        for (var i = targetItems.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapItem = targetItems[i];
            targetItems[i] = targetItems[j];
            targetItems[j] = swapItem;
        }
    }

    /**
     * X 座標または Y 座標（上端）で並べ替えて重ね順を組み直す
     * @param {PageItem[]} docSelection - 選択オブジェクト（2つ以上）
     * @param {string} axis - "x" / "y"
     * @param {string} order - "asc" / "desc" / "rand"
     * @returns {void}
     */
    function sortByAxis(docSelection, axis, order) {
        var targetItems = copyItems(docSelection);
        if (order === "rand") {
            shuffleItems(targetItems);
        } else {
            var boundsIndex = (axis === "x") ? 0 : 1;
            /* Y は上ほど値が大きいので、X と昇順・降順を逆にする / Y grows upward, so flip the direction for Y */
            var numericAscending = (axis === "y") ? (order === "desc") : (order !== "desc");
            targetItems.sort(function (a, b) {
                var aValue = a.geometricBounds[boundsIndex];
                var bValue = b.geometricBounds[boundsIndex];
                return numericAscending ? aValue - bValue : bValue - aValue;
            });
        }
        reorderItems(targetItems);
    }

    /**
     * 有効な選択を返す。ドキュメントが無い、または2つ未満なら知らせて null
     * @returns {PageItem[]|null} 選択オブジェクト
     */
    function getValidSelection() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return null;
        }
        var docSelection = app.activeDocument.selection;
        if (!docSelection || docSelection.length < 2) {
            alert(getLabel("alert.selectMore"));
            return null;
        }
        return docSelection;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンのパネルを作る
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {string[]} radioKeys - radio / tooltip の LABELS キー
     * @returns {RadioButton[]} 作ったラジオボタン（先頭を選択済み）
     */
    function addRadioPanel(parentWindow, titlePath, radioKeys) {
        var radioPanel = parentWindow.add("panel", undefined, getLabel(titlePath));
        radioPanel.orientation = "column";
        radioPanel.alignChildren = "left";
        radioPanel.margins = [15, 20, 15, 10];
        var panelRadios = [];
        for (var i = 0; i < radioKeys.length; i++) {
            var panelRadio = radioPanel.add("radiobutton", undefined, getLabel("radio." + radioKeys[i]));
            panelRadio.helpTip = getLabel("tooltip." + radioKeys[i]);
            panelRadios.push(panelRadio);
        }
        panelRadios[0].value = true;
        return panelRadios;
    }

    /**
     * ダイアログを開いたときに表示位置をずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 右へずらす量
     * @param {number} offsetY - 下へずらす量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを出し、ラジオボタンを押すたびに重ね順を組み直す。キャンセルで元の順へ戻す
     * @returns {void}
     */
    function main() {
        var sortDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        sortDialog.alignChildren = "left";

        var originalItems = copyItems(app.activeDocument.selection);

        var methodRadios = addRadioPanel(sortDialog, "panel.sortMethod", ["zOrder", "xAxis", "yAxis"]);
        var zOrderRadio = methodRadios[0];
        var xAxisRadio = methodRadios[1];
        var yAxisRadio = methodRadios[2];

        var orderRadios = addRadioPanel(sortDialog, "panel.orderMethod", ["asc", "desc", "random"]);
        var descRadio = orderRadios[1];
        var randomRadio = orderRadios[2];

        var btnRowGroup = sortDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "bottom"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOk = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 選んだ基準と並び順で重ね順を組み直す（ラジオボタンを押すたびに呼ぶ）
         * @returns {void}
         */
        function previewSort() {
            var order = "asc";
            if (randomRadio.value) {
                order = "rand";
            } else if (descRadio.value) {
                order = "desc";
            }

            var docSelection = getValidSelection();
            if (!docSelection) return;

            if (zOrderRadio.value) {
                /* 現在の重ね順：押すたびに逆順にする / current order: reverse on every click */
                reorderItems(copyItems(docSelection));
            } else if (xAxisRadio.value) {
                sortByAxis(docSelection, "x", order);
            } else if (yAxisRadio.value) {
                sortByAxis(docSelection, "y", order);
            }
            app.redraw();
        }

        var allRadios = methodRadios.concat(orderRadios);
        for (var i = 0; i < allRadios.length; i++) {
            allRadios[i].onClick = previewSort;
        }

        btnOk.onClick = function () {
            sortDialog.close(1);
        };
        btnCancel.onClick = function () {
            if (originalItems.length > 0) {
                reorderItems(originalItems);
                app.redraw();
            }
            sortDialog.close(0);
        };

        sortDialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(sortDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        sortDialog.show();
    }

    main();

})();
