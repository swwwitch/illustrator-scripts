#target illustrator
#targetengine "SortItemsByPositionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ページアイテムをX座標またはY座標で並べ替え、その順序で重ね順を更新します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortItemsByPosition.md

### Overview

Sorts page items by their X or Y coordinate and rewrites the stacking order to match.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortItemsByPosition.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortItemsByPosition";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortItemsByPosition.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortItemsByPosition.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var OPTION_PANEL_WIDTH = 200; /* 左カラムのパネル幅 / width of the left-column panels */
    var BUTTON_WIDTH = 100;       /* 右カラムのボタン幅 / width of the right-column buttons */

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title: { ja: "重ね順の変更", en: "Reorder Objects" }
        },
        panel: {
            sortKey: { ja: "ソート基準", en: "Sort Criteria" },
            target: { ja: "対象", en: "Target" }
        },
        radio: {
            xLeft: { ja: "X座標（左右）", en: "X (Horizontal)" },
            yTop: { ja: "Y座標（上下）", en: "Y (Vertical)" },
            random: { ja: "ランダム", en: "Random" },
            selection: { ja: "選択オブジェクト", en: "Selection" },
            artboard: { ja: "現在のアートボードに限定", en: "Limit to Current Artboard" }
        },
        checkbox: {
            moveLayer: { ja: "レイヤーを移動", en: "Move to Top Layer" }
        },
        tooltip: {
            xLeft: {
                ja: "左にあるものほど背面になるよう、X座標で重ね順を組み直します。",
                en: "Restacks by X, putting objects further left further back."
            },
            yTop: {
                ja: "上にあるものほど背面になるよう、Y座標で重ね順を組み直します。",
                en: "Restacks by Y, putting objects further up further back."
            },
            random: { ja: "重ね順をランダムに組み直します。", en: "Restacks the objects in random order." },
            selection: { ja: "選択しているオブジェクトだけを対象にします。", en: "Works on the selected objects only." },
            artboard: {
                ja: "現在のアートボードに載っているオブジェクトを対象にします。",
                en: "Works on the objects on the current artboard."
            },
            moveLayer: { ja: "対象オブジェクトを最前面のレイヤーへまとめて移します。", en: "Moves the objects onto the topmost layer." },
            reverse: { ja: "いまの重ね順をそのまま逆にします。", en: "Simply reverses the current stacking order." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            reverse: { ja: "反転", en: "Reverse" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * オブジェクトの左上がアートボードの内側にあるかを調べる
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {number[]} artboardBounds - [左, 上, 右, 下]
     * @returns {boolean} 内側なら true
     */
    function isInsideArtboard(pageItem, artboardBounds) {
        var itemLeft = pageItem.position[0];
        var itemTop = pageItem.position[1];
        return (itemLeft >= artboardBounds[0] && itemLeft <= artboardBounds[2] &&
            itemTop <= artboardBounds[1] && itemTop >= artboardBounds[3]);
    }

    /**
     * 対象オブジェクトを集める。選択があれば選択を優先する
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} useSelection - 選択オブジェクトを対象にする
     * @param {boolean} useArtboard - 現在のアートボード上のオブジェクトに限る（選択を使わないとき）
     * @returns {PageItem[]} 対象オブジェクト
     */
    function collectItems(doc, useSelection, useArtboard) {
        var targetItems = [];
        var i;
        if (useSelection && doc.selection.length > 0) {
            for (i = 0; i < doc.selection.length; i++) {
                targetItems.push(doc.selection[i]);
            }
            return targetItems;
        }

        if (useArtboard) {
            var artboardBounds = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            for (i = 0; i < doc.pageItems.length; i++) {
                if (isInsideArtboard(doc.pageItems[i], artboardBounds)) {
                    targetItems.push(doc.pageItems[i]);
                }
            }
        } else {
            for (i = 0; i < doc.pageItems.length; i++) {
                targetItems.push(doc.pageItems[i]);
            }
        }
        return targetItems;
    }

    /**
     * 選択範囲全体が縦長なら Y、横長なら X を初期の基準にする
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {string} "y" または "x"（2つ未満なら "x"）
     */
    function detectSortAxis(selectedItems) {
        if (selectedItems.length <= 1) return "x";
        var firstBounds = selectedItems[0].geometricBounds;
        var leftMost = firstBounds[0];
        var rightMost = firstBounds[2];
        var topMost = firstBounds[1];
        var bottomMost = firstBounds[3];
        for (var i = 1; i < selectedItems.length; i++) {
            var itemBounds = selectedItems[i].geometricBounds;
            if (itemBounds[0] < leftMost) leftMost = itemBounds[0];
            if (itemBounds[2] > rightMost) rightMost = itemBounds[2];
            if (itemBounds[1] > topMost) topMost = itemBounds[1];
            if (itemBounds[3] < bottomMost) bottomMost = itemBounds[3];
        }
        return (topMost - bottomMost >= rightMost - leftMost) ? "y" : "x";
    }

    /**
     * 対象オブジェクトが2つ以上のレイヤー（名前で区別）にまたがっているかを調べる
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {boolean} またがっていれば true
     */
    function spansMultipleLayers(targetItems) {
        var layerNames = {};
        for (var i = 0; i < targetItems.length; i++) {
            layerNames[targetItems[i].layer.name] = true;
        }
        var layerCount = 0;
        for (var layerName in layerNames) {
            layerCount++;
        }
        return layerCount > 1;
    }

    // =========================================
    // 並べ替えと重ね順 / Sorting and stacking
    // =========================================

    /**
     * 配列の順序をシャッフルする（直接書き換える）
     * @param {PageItem[]} targetItems - 対象の配列
     * @returns {void}
     */
    function shuffleItems(targetItems) {
        var remainingCount = targetItems.length;
        while (remainingCount) {
            var randomIndex = Math.floor(Math.random() * remainingCount--);
            var swapItem = targetItems[remainingCount];
            targetItems[remainingCount] = targetItems[randomIndex];
            targetItems[randomIndex] = swapItem;
        }
    }

    /**
     * 並べ替えの基準に従って配列を並べる（先頭ほど背面になる）
     * @param {PageItem[]} targetItems - 対象の配列（直接並べ替える）
     * @param {string} sortMode - "xAsc"（左ほど背面）/ "yDesc"（上ほど背面）/ "random"
     * @returns {void}
     */
    function sortItems(targetItems, sortMode) {
        if (sortMode === "yDesc") {
            targetItems.sort(function (a, b) { return b.position[1] - a.position[1]; });
        } else if (sortMode === "random") {
            shuffleItems(targetItems);
        } else {
            targetItems.sort(function (a, b) { return a.position[0] - b.position[0]; });
        }
    }

    /**
     * 配列の順に最前面へ送り、重ね順を配列の順（先頭が背面）にする
     * @param {PageItem[]} targetItems - 対象の配列
     * @returns {void}
     */
    function restackItems(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].zOrder(ZOrderMethod.BRINGTOFRONT);
        }
        app.redraw();
    }

    /**
     * 対象オブジェクトを最前面のレイヤーへ移し、オブジェクトの無くなったレイヤーを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 並べ替え済みの対象オブジェクト（逆順に並べ直す）
     * @returns {void}
     */
    function moveItemsToTopLayer(doc, targetItems) {
        var topLayer = doc.layers[0];
        /* 移動前に逆順にして順序を保つ / reverse first to keep the order */
        targetItems.reverse();
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].move(topLayer, ElementPlacement.PLACEATEND);
        }
        /* 移動後に重ね順を更新して見た目を保つ / restack after moving to keep the look */
        restackItems(targetItems);
        /* オブジェクトの無いレイヤーを削除 / remove layers left without page items */
        for (var j = doc.layers.length - 1; j >= 0; j--) {
            if (doc.layers[j].pageItems.length === 0) {
                doc.layers[j].remove();
            }
        }
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オプションのパネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titlePath - パネル名の LABELS パス
     * @returns {Panel} 作ったパネル
     */
    function addOptionPanel(parentGroup, titlePath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = "left";
        optionPanel.margins = [15, 20, 15, 10];
        optionPanel.preferredSize.width = OPTION_PANEL_WIDTH;
        return optionPanel;
    }

    /**
     * ラベルと tooltip の付いたコントロールを追加する
     * @param {Object} parentGroup - 追加先のグループまたはパネル
     * @param {string} controlType - "radiobutton" / "checkbox" / "button"
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} [tooltipPath] - tooltip の LABELS パス
     * @param {Object} [creationProperties] - add() に渡す作成時プロパティ
     * @returns {Object} 作ったコントロール
     */
    function addLabeledControl(parentGroup, controlType, labelPath, tooltipPath, creationProperties) {
        var labeledControl = creationProperties
            ? parentGroup.add(controlType, undefined, getLabel(labelPath), creationProperties)
            : parentGroup.add(controlType, undefined, getLabel(labelPath));
        if (tooltipPath) labeledControl.helpTip = getLabel(tooltipPath);
        return labeledControl;
    }

    /**
     * 重ね順を変更するダイアログを出す。基準や対象を切り替えるとその場で重ね順を組み直す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} initialItems - 最初の対象オブジェクト
     * @param {string} initialAxis - 最初に選ぶ基準（"x" / "y"）
     * @param {boolean} hasSelection - 選択があるか
     * @returns {Object|null} { sortMode, useSelection, useArtboard, moveToTopLayer }。キャンセル時は null
     */
    function showReorderDialog(doc, initialItems, initialAxis, hasSelection) {
        var previewItems = initialItems;

        var reorderDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        reorderDialog.orientation = "column";
        reorderDialog.alignChildren = "left";

        /* 2カラムレイアウト / two-column layout */
        var columnsGroup = reorderDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        var optionsColumn = columnsGroup.add("group");
        optionsColumn.orientation = "column";
        optionsColumn.alignChildren = "left";

        var sortKeyPanel = addOptionPanel(optionsColumn, "panel.sortKey");
        var xLeftRadio = addLabeledControl(sortKeyPanel, "radiobutton", "radio.xLeft", "tooltip.xLeft");
        var yTopRadio = addLabeledControl(sortKeyPanel, "radiobutton", "radio.yTop", "tooltip.yTop");
        var randomRadio = addLabeledControl(sortKeyPanel, "radiobutton", "radio.random", "tooltip.random");
        if (initialAxis === "y") {
            yTopRadio.value = true;
        } else {
            xLeftRadio.value = true;
        }

        var targetPanel = addOptionPanel(optionsColumn, "panel.target");
        var selectionRadio = addLabeledControl(targetPanel, "radiobutton", "radio.selection", "tooltip.selection");
        var artboardRadio = addLabeledControl(targetPanel, "radiobutton", "radio.artboard", "tooltip.artboard");
        selectionRadio.value = hasSelection;
        artboardRadio.value = !hasSelection;

        var moveLayerGroup = optionsColumn.add("group");
        moveLayerGroup.orientation = "column";
        moveLayerGroup.alignChildren = "center";
        moveLayerGroup.preferredSize.width = OPTION_PANEL_WIDTH;
        var moveTopLayerCheckbox = addLabeledControl(moveLayerGroup, "checkbox", "checkbox.moveLayer", "tooltip.moveLayer");
        moveTopLayerCheckbox.value = false;
        /* 対象が1つのレイヤーだけなら移動は不要 / nothing to move when all targets share one layer */
        if (!spansMultipleLayers(initialItems)) {
            moveTopLayerCheckbox.enabled = false;
        }

        var btnColumn = columnsGroup.add("group");
        btnColumn.orientation = "column";
        btnColumn.alignChildren = "right";
        var btnOk = addLabeledControl(btnColumn, "button", "button.ok", null, { name: "OK" });
        btnOk.preferredSize.width = BUTTON_WIDTH;
        var btnReverse = addLabeledControl(btnColumn, "button", "button.reverse", "tooltip.reverse");
        btnReverse.preferredSize.width = BUTTON_WIDTH;
        var spacer = btnColumn.add("statictext", undefined, "");
        spacer.preferredSize.height = 100;
        var btnCancel = addLabeledControl(btnColumn, "button", "button.cancel");
        btnCancel.preferredSize.width = BUTTON_WIDTH;

        /**
         * 選ばれている並べ替えの基準を返す
         * @returns {string} "xAsc" / "yDesc" / "random"
         */
        function getSortMode() {
            if (yTopRadio.value) return "yDesc";
            if (randomRadio.value) return "random";
            return "xAsc";
        }

        /**
         * 並べ替えて重ね順を組み直す（プレビュー）
         * @param {string} sortMode - 並べ替えの基準
         * @returns {void}
         */
        function previewSort(sortMode) {
            sortItems(previewItems, sortMode);
            restackItems(previewItems);
        }

        xLeftRadio.onClick = function () {
            previewSort("xAsc");
        };
        yTopRadio.onClick = function () {
            previewSort("yDesc");
        };
        randomRadio.onClick = function () {
            previewSort("random");
        };
        selectionRadio.onClick = function () {
            previewItems = collectItems(doc, selectionRadio.value === true, artboardRadio.value === true);
            previewSort(getSortMode());
        };
        artboardRadio.onClick = selectionRadio.onClick;
        btnReverse.onClick = function () {
            previewItems.reverse();
            restackItems(previewItems);
        };

        prepareDialogWindow(reorderDialog, SCRIPT_NAME);
        if (reorderDialog.show() != 1) return null;
        return {
            sortMode: getSortMode(),
            useSelection: selectionRadio.value === true,
            useArtboard: artboardRadio.value === true,
            moveToTopLayer: moveTopLayerCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで選んだ基準で重ね順を組み直し、必要なら最前面のレイヤーへ移す
     * @returns {void}
     */
    function main() {
        var doc = app.activeDocument;
        var hasSelection = (doc.selection.length > 0);
        var initialAxis = hasSelection ? detectSortAxis(doc.selection) : "x";
        var initialItems = collectItems(doc, hasSelection, !hasSelection);

        var reorderOptions = showReorderDialog(doc, initialItems, initialAxis, hasSelection);
        if (!reorderOptions) return;

        var targetItems = collectItems(doc, reorderOptions.useSelection, reorderOptions.useArtboard);
        sortItems(targetItems, reorderOptions.sortMode);
        restackItems(targetItems);
        if (reorderOptions.moveToTopLayer) {
            moveItemsToTopLayer(doc, targetItems);
        }
    }

    main();

})();
