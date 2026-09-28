#target illustrator
#targetengine "SmartSelectionFilterSimpleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを条件に応じてフィルタリングし、テキスト／オープンパス／クローズパスを選択し直します。
対象スコープを切り替えると、選択直下だけでなくグループ内のオブジェクトも対象にできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSelectionFilterSimple.md

### Overview

Filters the selection by condition and reselects text frames, open paths, or closed paths.
Switching the scope extends the filter from the top level of the selection to objects inside groups as well.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSelectionFilterSimple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartSelectionFilterSimple";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSelectionFilterSimple.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSelectionFilterSimple.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "選択フィルター", en: "Selection Filter" },
        panelCondition: { ja: "条件", en: "Conditions" },
        targetScopePanel: { ja: "対象スコープ", en: "Selection Scope" },
        text: { ja: "テキスト", en: "Text" },
        openPath: { ja: "オープンパス", en: "Open Path" },
        closePath: { ja: "クローズパス", en: "Closed Path" },
        scopeSelectedOnly: { ja: "選択直下のみ", en: "Selected objects only" },
        scopeIncludeGroupItems: { ja: "グループ内も対象に含める", en: "Include objects inside groups" },
        scopeHint: { ja: "※ グループ内のオブジェクトを対象に含めるかを指定します", en: "* Choose whether to include objects inside groups" },
        btnOK: { ja: "OK", en: "OK" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        errNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        errNoSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        errNoCheck: { ja: "少なくとも1つチェックを入れてください。", en: "Please select at least one option." },
        errApply: { ja: "選択の更新に失敗しました。", en: "Failed to update selection." },
        errGeneral: { ja: "エラーが発生しました", en: "An error occurred" }
    };

    function getLabel(labelKey) {
        return LABELS[labelKey] ? LABELS[labelKey][currentLanguage] : labelKey;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // フィルタリング選択スクリプト
    (function () {
        main();

        function main() {
            var documentRef = null;
            var originalSelectionItems = [];
            var shouldRestoreOriginalSelection = true;

            function getOriginalSelectionItems(targetDocument) {
                var snapshotItems = [];

                if (!targetDocument.selection || targetDocument.selection.length === 0) {
                    return snapshotItems;
                }

                for (var selectionIndex = 0; selectionIndex < targetDocument.selection.length; selectionIndex++) {
                    snapshotItems.push(targetDocument.selection[selectionIndex]);
                }

                return snapshotItems;
            }

            function restoreOriginalSelection(targetDocument, originalItems) {
                if (!targetDocument || !originalItems) {
                    return false;
                }

                try {
                    targetDocument.selection = originalItems;
                    return true;
                } catch (restoreError) {
                    return false;
                }
            }

            try {
                if (app.documents.length === 0) {
                    alert(getLabel('errNoDoc'));
                    return;
                }

                documentRef = app.activeDocument;
                originalSelectionItems = getOriginalSelectionItems(documentRef);
                var expandedGroupTargetItemsCache = null;

                function applySelectionSafely(targetDocument, targetItems) {
                    if (!targetDocument || !targetItems) {
                        return false;
                    }

                    try {
                        targetDocument.selection = null;

                        for (var targetIndex = 0; targetIndex < targetItems.length; targetIndex++) {
                            try {
                                targetItems[targetIndex].selected = true;
                            } catch (itemSelectionError) {
                                // 選択できない項目はスキップ / Skip items that cannot be selected
                            }
                        }

                        return true;
                    } catch (selectionError) {
                        return false;
                    }
                }

                /* 選択対象候補を収集 / Collect selectable candidate items */

                function collectSelectableCandidateItems(sourceItems) {
                    var candidateItems = [];

                    for (var sourceIndex = 0; sourceIndex < sourceItems.length; sourceIndex++) {
                        collectSelectableCandidateItemsFromItem(sourceItems[sourceIndex], candidateItems);
                    }

                    return candidateItems;
                }

                /* スコープに応じて候補を取得 / Get candidate items based on scope */

                function getCandidateItemsForScope(includeGroupItems) {
                    if (!includeGroupItems) {
                        return originalSelectionItems;
                    }

                    if (!expandedGroupTargetItemsCache) {
                        expandedGroupTargetItemsCache = collectSelectableCandidateItems(originalSelectionItems);
                    }

                    return expandedGroupTargetItemsCache;
                }

                function collectSelectableCandidateItemsFromItem(candidateItem, candidateItems) {
                    if (!candidateItem) {
                        return;
                    }

                    /* グループ内のオブジェクトを再帰的に収集 / Recursively collect items inside groups */
                    /* クリップグループではマスク用パス（clipping path）を除外 / Exclude clipping mask paths in clipped groups */
                    /* 注意：クリップ範囲外の見た目までは考慮しない（内部構造ベースで判定） / Note: Does not evaluate visual clipping bounds (structure-based only) */
                    if (candidateItem.typename === "GroupItem") {
                        for (var childIndex = 0; childIndex < candidateItem.pageItems.length; childIndex++) {
                            var childItem = candidateItem.pageItems[childIndex];
                            if (candidateItem.clipped && isClippingPathItem(childItem)) {
                                /* マスク用パス自体は選択対象から除外 / Exclude the clipping mask path itself */
                                continue;
                            }

                            collectSelectableCandidateItemsFromItem(childItem, candidateItems);
                        }
                        return;
                    }

                    /* CompoundPathItem：親オブジェクトとして扱う / Treat as parent object */
                    if (candidateItem.typename === "CompoundPathItem") {
                        candidateItems.push(candidateItem);
                        return;
                    }

                    /* TODO: SymbolItem 対応 / TODO: Support SymbolItem */

                    candidateItems.push(candidateItem);
                }

                function isClippingPathItem(item) {
                    if (!item) {
                        return false;
                    }

                    try {
                        if (item.typename === "PathItem" && item.clipping) {
                            return true;
                        }

                        if (item.typename === "CompoundPathItem") {
                            for (var pathIndex = 0; pathIndex < item.pathItems.length; pathIndex++) {
                                if (item.pathItems[pathIndex].clipping) {
                                    return true;
                                }
                            }
                        }
                    } catch (e) {
                        return false;
                    }

                    return false;
                }

                function hasVisibleStroke(item) {
                    try {
                        if (!item.stroked) return false;
                        if (!item.strokeColor || item.strokeColor.typename === "NoColor") return false;
                        if (item.strokeWidth <= 0) return false;
                        if (item.opacity !== undefined && item.opacity === 0) return false;
                        return true;
                    } catch (e) {
                        return false;
                    }
                }

                function hasVisibleFill(item) {
                    try {
                        if (!item.filled) return false;
                        if (!item.fillColor || item.fillColor.typename === "NoColor") return false;
                        if (item.opacity !== undefined && item.opacity === 0) return false;
                        return true;
                    } catch (e) {
                        return false;
                    }
                }

                function isSelectableItem(item) {
                    if (!item) {
                        return false;
                    }

                    try {
                        if (item.locked || item.hidden) {
                            return false;
                        }
                    } catch (e) {
                        return false;
                    }

                    return isParentChainSelectable(item);
                }

                function isParentChainSelectable(item) {
                    var parentItem = item.parent;

                    while (parentItem) {
                        try {
                            if (parentItem.typename === "Layer") {
                                if (parentItem.locked || !parentItem.visible) {
                                    return false;
                                }
                            } else if (parentItem.typename === "GroupItem") {
                                if (parentItem.locked || parentItem.hidden) {
                                    return false;
                                }
                            }
                        } catch (e) {
                            return false;
                        }

                        if (parentItem.typename === "Document") {
                            break;
                        }

                        parentItem = parentItem.parent;
                    }

                    return true;
                }

                function applyDefaultDialogValues(dialogUi) {
                    dialogUi.textCheckbox.value = true;
                    dialogUi.openPathCheckbox.value = false;
                    dialogUi.closedPathCheckbox.value = false;
                    dialogUi.scopeIncludeGroupItemsRadio.value = true;
                    dialogUi.scopeSelectedOnlyRadio.value = false;
                }

                function bindExclusiveOptionClick(targetCheckbox, allConditionCheckboxes) {
                    targetCheckbox.onClick = function () {
                        var keyboardState = ScriptUI.environment.keyboardState;

                        if (!keyboardState || !keyboardState.altKey) {
                            return;
                        }

                        for (var checkboxIndex = 0; checkboxIndex < allConditionCheckboxes.length; checkboxIndex++) {
                            allConditionCheckboxes[checkboxIndex].value = false;
                        }

                        targetCheckbox.value = true;
                    };
                }

                function createConditionPanel(parent) {
                    /* 条件パネル / Conditions panel */
                    var panel = parent.add("panel", undefined, getLabel('panelCondition'));
                    panel.orientation = "column";
                    panel.alignChildren = "left";
                    panel.margins = [15, 20, 15, 10];

                    var textCheckbox = panel.add("checkbox", undefined, getLabel('text'));
                    var openPathCheckbox = panel.add("checkbox", undefined, getLabel('openPath'));
                    var closedPathCheckbox = panel.add("checkbox", undefined, getLabel('closePath'));

                    var conditionCheckboxes = [textCheckbox, openPathCheckbox, closedPathCheckbox];
                    bindExclusiveOptionClick(textCheckbox, conditionCheckboxes);
                    bindExclusiveOptionClick(openPathCheckbox, conditionCheckboxes);
                    bindExclusiveOptionClick(closedPathCheckbox, conditionCheckboxes);

                    return {
                        textCheckbox: textCheckbox,
                        openPathCheckbox: openPathCheckbox,
                        closedPathCheckbox: closedPathCheckbox
                    };
                }

                function createScopePanel(parent) {
                    /* 対象スコープパネル / Selection scope panel */
                    var panel = parent.add("panel", undefined, getLabel('targetScopePanel'));
                    panel.orientation = "column";
                    panel.alignChildren = "left";
                    panel.margins = [15, 20, 15, 10];

                    var scopeSelectedOnlyRadio = panel.add("radiobutton", undefined, getLabel('scopeSelectedOnly'));
                    var scopeIncludeGroupItemsRadio = panel.add("radiobutton", undefined, getLabel('scopeIncludeGroupItems'));

                    /* 対象スコープのツールチップ / Tooltip for selection scope */
                    panel.helpTip = getLabel('scopeHint');
                    scopeSelectedOnlyRadio.helpTip = getLabel('scopeHint');
                    scopeIncludeGroupItemsRadio.helpTip = getLabel('scopeHint');

                    return {
                        scopeSelectedOnlyRadio: scopeSelectedOnlyRadio,
                        scopeIncludeGroupItemsRadio: scopeIncludeGroupItemsRadio
                    };
                }

                function createButtonGroup(parent) {
                    /* ボタン群 / Button group */
                    var group = parent.add("group");
                    group.orientation = "row";
                    group.alignment = "center";
                    group.margins = [0, 10, 0, 0];

                    var cancelButton = group.add("button", undefined, getLabel('btnCancel'), { name: "cancel" });
                    var okButton = group.add("button", undefined, getLabel('btnOK'), { name: "ok" });
                    cancelButton.preferredSize.width = 80;
                    okButton.preferredSize.width = 80;

                    return {
                        cancelButton: cancelButton,
                        okButton: okButton
                    };
                }

                function createDialog() {
                    var dialog = new Window("dialog", getLabel('dialogTitle') + " " + SCRIPT_VERSION);
                    dialog.orientation = "column";
                    dialog.alignChildren = "left";
                    dialog.margins = 20;

                    var conditionUi = createConditionPanel(dialog);
                    var scopeUi = createScopePanel(dialog);
                    var buttonUi = createButtonGroup(dialog);

                    var dialogUi = {
                        dialog: dialog,
                        textCheckbox: conditionUi.textCheckbox,
                        openPathCheckbox: conditionUi.openPathCheckbox,
                        closedPathCheckbox: conditionUi.closedPathCheckbox,
                        scopeSelectedOnlyRadio: scopeUi.scopeSelectedOnlyRadio,
                        scopeIncludeGroupItemsRadio: scopeUi.scopeIncludeGroupItemsRadio,
                        cancelButton: buttonUi.cancelButton,
                        okButton: buttonUi.okButton
                    };

                    applyDefaultDialogValues(dialogUi);

                    return dialogUi;
                }

                // =========================================
                // ダイアログ作成 / Create dialog
                // =========================================
                var dialogUi = createDialog();
                var dialog = dialogUi.dialog;

                function readDialogOptions(dialogUi) {
                    return {
                        includeText: dialogUi.textCheckbox.value,
                        includeOpenPath: dialogUi.openPathCheckbox.value,
                        includeClosePath: dialogUi.closedPathCheckbox.value,
                        includeGroupItems: dialogUi.scopeIncludeGroupItemsRadio.value
                    };
                }

                /* フィルター条件に一致するか判定 / Check if item matches filter conditions */

                function matchesFilterConditions(candidateItem, options) {
                    /* テキスト / Text */
                    if (options.includeText && candidateItem.typename === "TextFrame") {
                        return true;
                    }

                    /* パスアイテム / Path item */
                    if (candidateItem.typename === "PathItem") {
                        /* オープンパス / Open path */
                        if (options.includeOpenPath && !candidateItem.closed) {
                            return true;
                        }
                        /* クローズパス / Closed path */
                        if (options.includeClosePath && candidateItem.closed && (hasVisibleStroke(candidateItem) || hasVisibleFill(candidateItem))) {
                            return true;
                        }
                    }

                    /* 複合パス / Compound path */
                    if (candidateItem.typename === "CompoundPathItem") {
                        if (options.includeClosePath && (hasVisibleStroke(candidateItem) || hasVisibleFill(candidateItem))) {
                            return true;
                        }
                    }

                    return false;
                }

                function hasAnyConditionEnabled(options) {
                    return options.includeText || options.includeOpenPath || options.includeClosePath;
                }

                /* 条件一致かつ選択可能なオブジェクトを抽出 / Extract selectable items matching conditions */

                function getMatchedSelectableItems(candidateItems, options) {
                    var matchedSelectableItems = [];

                    for (var candidateIndex = 0; candidateIndex < candidateItems.length; candidateIndex++) {
                        var candidateItem = candidateItems[candidateIndex];

                        if (!isSelectableItem(candidateItem)) continue;
                        if (!matchesFilterConditions(candidateItem, options)) continue;

                        matchedSelectableItems.push(candidateItem);
                    }

                    return matchedSelectableItems;
                }

                /* フィルターを適用して選択更新 / Apply filter and update selection */

                function applyFilterSelection(options) {
                    if (!hasAnyConditionEnabled(options)) {
                        alert(getLabel('errNoCheck'));
                        return false;
                    }

                    var candidateItemsForScope = getCandidateItemsForScope(options.includeGroupItems);
                    var matchedSelectableItems = getMatchedSelectableItems(candidateItemsForScope, options);

                    if (!applySelectionSafely(documentRef, matchedSelectableItems)) {
                        alert(getLabel('errApply'));
                        return false;
                    }

                    return true;
                }

                // =========================================
                // OKボタン処理 / OK button handler
                // =========================================
                dialogUi.okButton.onClick = function () {
                    var options = readDialogOptions(dialogUi);

                    if (!applyFilterSelection(options)) {
                        return;
                    }

                    shouldRestoreOriginalSelection = false;
                    dialog.close();
                };

                dialogUi.cancelButton.onClick = function () {
                    dialog.close();
                };

                prepareDialogWindow(dialog, SCRIPT_NAME);
                dialog.show();

            } catch (error) {
                var lineInfo = (error && error.line) ? ("\nLine: " + error.line) : "";
                var message = getLabel('errGeneral') + ":\n" + error + lineInfo;

                try {
                    $.writeln(message);
                } catch (logError) { }

                alert(message);
            } finally {
                if (shouldRestoreOriginalSelection && documentRef && originalSelectionItems.length > 0) {
                    restoreOriginalSelection(documentRef, originalSelectionItems);
                }
            }
        }

    })();

})();
