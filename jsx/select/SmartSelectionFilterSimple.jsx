#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択オブジェクトを条件に応じてフィルタリングし、
テキスト／オープンパス／クローズパスを選択し直すスクリプト。
対象スコープを切り替えることで、選択直下のみ、またはグループ内のオブジェクトも対象にできます。
クローズパスは、塗りのみ／線のみ／塗り＋線のいずれも対象にします。
複合パスは親オブジェクトとして扱い、クリップグループではマスク用パスを選択対象から除外します。
選択更新時は、ロック／非表示のオブジェクトや親階層を避けて安全に処理します。

Overview

Filters selected objects by condition and reselects text,
open paths, and closed paths.
The selection scope can be limited to selected objects only,
or include objects inside groups.
Closed paths include fill-only, stroke-only, and fill-plus-stroke paths.
Compound paths are treated as parent objects, and clipping mask paths
inside clipping groups are excluded from the selection targets.
Selection updates are applied safely by avoiding locked or hidden objects
and locked or hidden parent containers.
*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================
var SCRIPT_VERSION = "v1.0";

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

// フィルタリング選択スクリプト
(function () {
    main();

    function main() {
        var activeDocument = null;
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

            activeDocument = app.activeDocument;
            originalSelectionItems = getOriginalSelectionItems(activeDocument);
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

                if (!applySelectionSafely(activeDocument, matchedSelectableItems)) {
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

            dialog.show();

        } catch (error) {
            var lineInfo = (error && error.line) ? ("\nLine: " + error.line) : "";
            var message = getLabel('errGeneral') + ":\n" + error + lineInfo;

            try {
                $.writeln(message);
            } catch (logError) { }

            alert(message);
        } finally {
            if (shouldRestoreOriginalSelection && activeDocument && originalSelectionItems.length > 0) {
                restoreOriginalSelection(activeDocument, originalSelectionItems);
            }
        }
    }

})();