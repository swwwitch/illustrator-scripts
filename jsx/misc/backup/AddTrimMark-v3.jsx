#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**
 * 【Illustrator用スクリプト】
 * ダイアログで対象範囲とトンボ設定を選び、トリムマークを作成して「トンボ」レイヤーに配置するスクリプト
 *
 * 【概要】
 * ・実行時にダイアログボックスを表示し、トリムマーク作成の対象範囲を選択できます。
 * ・対象範囲は「選択オブジェクト」「現在のアートボード」「すべてのアートボード」から選べます。
 * ・オブジェクトを選択している場合は、「選択オブジェクト」が初期選択されます。
 * ・「日本式トンボ」チェックボックスで、日本式トンボ／欧米式トンボを切り替えできます。
 * ・選択オブジェクト対象では、最初に選択されているオブジェクトをその場で複製し、塗り・線をなしにしてトリムマークを作成します。
 * ・アートボード対象では、各アートボードと同じ位置・サイズの長方形を一時レイヤー上に作成してトリムマークを作成します。
 * ・「すべてのアートボード」では、各アートボードごとに個別にトリムマークを作成します。
 * ・作成後、一時オブジェクトのみを削除し、作成されたトリムマークを「トンボ」レイヤーへ移動したあとで一時レイヤーを後始末します。
 * ・「トンボ」レイヤーが存在しない場合は新規作成します。
 * ・トリムマーク作成前にトンボ種別の設定を一時的に変更し、全処理の終了後に元の設定へ戻します。
 * ・「既存の『トンボ』レイヤーを削除してから作成」を有効にした場合は、親レイヤーを含む必要なロック・表示・テンプレート状態を一時的に調整して削除し、処理後は状態を可能な限り復元します。
 * ・エラー時も、一時オブジェクト・一時レイヤー・環境設定を可能な限り後始末するようにしています。
 *
 * 【処理の流れ】
 * 1. ダイアログボックスで対象範囲とトンボ種別を選択
 * 2. 対象に応じて一時オブジェクトを作成または複製
 * 3. 一時オブジェクトの塗りと線をなしに設定
 * 4. 選択したトンボ種別を一時的に有効化し、対象ごとにトリムマーク作成を実行
 * 5. 作成されたトリムマークを取得し、一時オブジェクトのみを削除する
 * 6. 必要に応じて既存の「トンボ」レイヤーを削除する
 * 7. 「トンボ」レイヤーを取得または新規作成する
 * 8. 作成されたトリムマークを「トンボ」レイヤーに移動する
 * 9. 作成したトリムマークの選択を解除する
 * 10. 環境設定を元に戻し、「トンボ」レイヤーをロックして終了する
 *
 * [Illustrator Script]
 * A script that lets you choose the target range and crop-mark style in a dialog, then creates trim marks and places them on the "トンボ" layer.
 *
 * [Overview]
 * - Shows a dialog at launch so you can choose the target range for creating trim marks.
 * - The target range can be set to "Selected Objects", "Current Artboard", or "All Artboards".
 * - When objects are already selected, "Selected Objects" is chosen by default.
 * - The "Use Japanese Crop Marks" checkbox switches between Japanese-style and Western-style crop marks.
 * - For the selected-objects mode, the script duplicates the first selected object in place, removes fill and stroke, and creates trim marks from that duplicate.
 * - For artboard-based modes, the script creates temporary rectangles on a temporary layer to match each artboard and uses them to generate trim marks.
 * - In "All Artboards" mode, trim marks are created separately for each artboard.
 * - After creation, only the temporary source objects are removed first, then the created trim marks are moved to the "トンボ" layer, and finally the temporary layer is cleaned up.
 * - If the "トンボ" layer does not exist, it is created automatically.
 * - The crop-mark style preference is changed temporarily before execution and restored after all processing finishes.
 * - When "Remove existing \"トンボ\" layer before creating" is enabled, the script temporarily adjusts the required parent-layer lock, visibility, and template states, removes the layer, and restores those states as much as possible afterward.
 * - Even on error, the script attempts to clean up temporary objects, the temporary layer, and preferences as much as possible.
 *
 * [Process]
 * 1. Choose the target range and crop-mark style in the dialog.
 * 2. Create or duplicate temporary objects depending on the selected mode.
 * 3. Set the fill and stroke of temporary objects to none.
 * 4. Temporarily enable the selected crop-mark style and run trim-mark creation for each target.
 * 5. Collect the created trim marks, then remove only the temporary source objects.
 * 6. Remove the existing "トンボ" layer if needed.
 * 7. Find the "トンボ" layer, or create it if needed.
 * 8. Move the created trim marks to the "トンボ" layer.
 * 9. Clear the selection of the created trim marks.
 * 10. Restore preferences, lock the "トンボ" layer, and finish.
 *
 * 【作成日 / Created】2025-02-05
 * 【更新日 / Updated】2026-04-01
 */

// =========================================
// バージョンとローカライズ
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "トンボを作成",
        en: "Create Trim Marks"
    },
    panelTarget: {
        ja: "作成対象",
        en: "Creation Target"
    },
    panelOption: {
        ja: "オプション",
        en: "Options"
    },
    rbSelection: {
        ja: "選択オブジェクト",
        en: "Selected Objects"
    },
    rbActiveArtboard: {
        ja: "現在のアートボード",
        en: "Current Artboard"
    },
    rbAllArtboards: {
        ja: "すべてのアートボード",
        en: "All Artboards"
    },
    chkJapanese: {
        ja: "日本式トンボ",
        en: "Use Japanese Crop Marks"
    },
    chkReplaceLayer: {
        ja: "既存の「トンボ」レイヤーを削除してから作成",
        en: "Remove existing \"トンボ\" layer before creating"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnOk: {
        ja: "OK",
        en: "OK"
    },
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errorNoSelection: {
        ja: "オブジェクトが選択されていません。",
        en: "No objects are selected."
    },
    errorNoTarget: {
        ja: "トリムマーク作成対象がありません。",
        en: "There is no target for creating trim marks."
    },
    errorRemoveTrimLayer: {
        ja: '既存の「トンボ」レイヤーを削除できませんでした。',
        en: 'Failed to remove the existing "トンボ" layer.'
    },
};

function L(key) {
    return LABELS[key][lang];
}

function __logError(message, err) {
    try {
        $.writeln('[AddTrimMark] ' + message + (err ? ' : ' + err : ''));
    } catch (logError) {
    }
}

(function () {
    if (app.documents.length === 0) {
        alert(L('alertNoDocument'));
        return;
    }

    var doc = app.activeDocument;
    var TEMP_LAYER_NAME = '__AddTrimMark_Temp__';
    var tempItems = [];
    var trimLayer = null;
    var marks = [];
    var cropMarkStyleBefore = app.preferences.getBooleanPreference('cropMarkStyle');
    var tempLayer = null;
    var activeLayerBefore = doc.activeLayer;
    var adjustedLayerStates = [];

    function findAdjustedLayerState(layer, adjustedList) {
        for (var i = 0; i < adjustedList.length; i++) {
            if (adjustedList[i].layer === layer) {
                return adjustedList[i];
            }
        }
        return null;
    }

    function ensureLayerRemovable(layer, adjustedList) {
        var current = layer;
        while (current) {
            if (current.typename !== 'Layer') {
                break;
            }

            var state = findAdjustedLayerState(current, adjustedList);
            if (!state) {
                state = {
                    layer: current,
                    locked: current.locked,
                    visible: current.visible,
                    printable: (typeof current.printable !== 'undefined') ? current.printable : null,
                    template: (typeof current.template !== 'undefined') ? current.template : null
                };
                adjustedList.push(state);
            }

            try {
                if (current.locked) {
                    current.locked = false;
                }
            } catch (unlockError) {
                __logError('Failed to unlock layer', unlockError);
            }

            try {
                if (!current.visible) {
                    current.visible = true;
                }
            } catch (visibleError) {
                __logError('Failed to make layer visible', visibleError);
            }

            try {
                if (typeof current.template !== 'undefined' && current.template) {
                    current.template = false;
                }
            } catch (templateError) {
                __logError('Failed to disable template state', templateError);
            }

            current = current.parent;
            if (!current || !current.typename || current.typename !== 'Layer') {
                break;
            }
        }
    }

    function restoreAdjustedLayerStates(adjustedList) {
        for (var i = adjustedList.length - 1; i >= 0; i--) {
            var state = adjustedList[i];
            try {
                if (state.template !== null && typeof state.layer.template !== 'undefined') {
                    state.layer.template = state.template;
                }
            } catch (restoreTemplateError) {
                __logError('Failed to restore template state', restoreTemplateError);
            }
            try {
                if (state.printable !== null && typeof state.layer.printable !== 'undefined') {
                    state.layer.printable = state.printable;
                }
            } catch (restorePrintableError) {
                __logError('Failed to restore printable state', restorePrintableError);
            }
            try {
                state.layer.visible = state.visible;
            } catch (restoreVisibleError) {
                __logError('Failed to restore visibility', restoreVisibleError);
            }
            try {
                state.layer.locked = state.locked;
            } catch (restoreLockedError) {
                __logError('Failed to restore lock state', restoreLockedError);
            }
        }
        adjustedList.length = 0;
    }

    function removeExistingTrimLayers() {
        var failedCount = 0;
        for (var r = doc.layers.length - 1; r >= 0; r--) {
            if (doc.layers[r].name === 'トンボ') {
                try {
                    ensureLayerRemovable(doc.layers[r], adjustedLayerStates);
                    doc.layers[r].remove();
                } catch (removeLayerError) {
                    failedCount++;
                    __logError('Failed to remove existing トンボ layer', removeLayerError);
                }
            }
        }
        if (failedCount > 0) {
            throw new Error(L('errorRemoveTrimLayer'));
        }
    }

    function getOrCreateTempLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === TEMP_LAYER_NAME) {
                tempLayer = doc.layers[i];
                tempLayer.locked = false;
                tempLayer.visible = true;
                doc.activeLayer = tempLayer;
                return tempLayer;
            }
        }

        tempLayer = doc.layers.add();
        tempLayer.name = TEMP_LAYER_NAME;
        tempLayer.locked = false;
        tempLayer.visible = true;
        doc.activeLayer = tempLayer;
        return tempLayer;
    }



    function showTargetDialog() {
        function findTrimLayer() {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === 'トンボ') {
                    return doc.layers[i];
                }
            }
            return null;
        }
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = 'left';
        dlg.margins = 16;

        var pnl = dlg.add('panel', undefined, L('panelTarget'));
        pnl.orientation = 'column';
        pnl.alignChildren = 'left';
        pnl.margins = [15, 20, 15, 10];

        var rbSelection = pnl.add('radiobutton', undefined, L('rbSelection'));
        var rbActiveArtboard = pnl.add('radiobutton', undefined, L('rbActiveArtboard'));
        var rbAllArtboards = pnl.add('radiobutton', undefined, L('rbAllArtboards'));

        /* 選択状態に応じてデフォルトを切り替え / Switch default selection based on current selection */
        if (doc.selection.length > 0) {
            rbSelection.value = true;
        } else {
            rbActiveArtboard.value = true;
        }

        var pnlOption = dlg.add('panel', undefined, L('panelOption'));
        pnlOption.orientation = 'column';
        pnlOption.alignChildren = 'left';
        pnlOption.alignment = 'fill';
        pnlOption.margins = [15, 20, 15, 10];

        var chkJapanese = pnlOption.add('checkbox', undefined, L('chkJapanese'));
        chkJapanese.value = true;

        var chkReplaceLayer = pnlOption.add('checkbox', undefined, L('chkReplaceLayer'));

        var existingTrimLayer = findTrimLayer();
        if (existingTrimLayer) {
            chkReplaceLayer.value = true;
            chkReplaceLayer.enabled = true;
        } else {
            chkReplaceLayer.value = false;
            chkReplaceLayer.enabled = false;
        }

        var btns = dlg.add('group');
        btns.alignment = 'center';
        var btnCancel = btns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        var btnOk = btns.add('button', undefined, L('btnOk'), { name: 'ok' });

        if (dlg.show() !== 1) {
            return null;
        }

        var mode = rbSelection.value ? 'selection'
            : rbAllArtboards.value ? 'allArtboards'
                : 'activeArtboard';

        return {
            mode: mode,
            useJapaneseCrop: chkJapanese.value,
            replaceLayer: chkReplaceLayer.value
        };
    }


    function addTempItemFromArtboard(index) {
        var rect = doc.artboards[index].artboardRect;
        var left = rect[0];
        var top = rect[1];
        var right = rect[2];
        var bottom = rect[3];
        var width = right - left;
        var height = top - bottom;
        var targetLayer = getOrCreateTempLayer();
        doc.activeLayer = targetLayer;
        var item = targetLayer.pathItems.rectangle(top, left, width, height);
        item.filled = false;
        item.stroked = false;
        tempItems.push(item);
    }

    // =========================================
    // 後始末 / Cleanup
    // =========================================
    function removeTempItems() {
        for (var k = 0; k < tempItems.length; k++) {
            try {
                tempItems[k].remove();
            } catch (removeTempError) {
                __logError('Failed to remove temporary item', removeTempError);
            }
        }
        tempItems = [];
    }

    function cleanupTempLayer() {
        if (tempLayer) {
            try {
                if (tempLayer.pageItems.length === 0 && tempLayer.layers.length === 0) {
                    tempLayer.remove();
                }
            } catch (removeTempLayerError) {
                __logError('Failed to remove temporary layer', removeTempLayerError);
            }
            tempLayer = null;
        }
    }

    /* トリムマーク作成実行 / Execute trim mark creation */
    function executeTrimMarkForCurrentTempItems() {
        if (tempItems.length === 0) {
            throw new Error(L('errorNoTarget'));
        }

        if (tempLayer) {
            doc.activeLayer = getOrCreateTempLayer();
        }

        doc.selection = null;
        doc.selection = tempItems.slice(0);
        app.redraw();
        app.executeMenuCommand('TrimMark v25');

        var createdMarks = doc.selection.slice(0);
        removeTempItems();
        return createdMarks;
    }

    // =========================================
    // ダイアログと実行 / Dialog and execution
    // =========================================
    var dialogResult = showTargetDialog();
    if (!dialogResult) {
        return;
    }

    var targetMode = dialogResult.mode;
    var useJapaneseCrop = dialogResult.useJapaneseCrop;
    var replaceLayer = dialogResult.replaceLayer;

    try {
        app.preferences.setBooleanPreference('cropMarkStyle', useJapaneseCrop);

        if (targetMode === 'selection') {
            var sourceSelection = doc.selection.slice(0);
            marks = [];

            if (sourceSelection.length === 0) {
                throw new Error(L('errorNoSelection'));
            }

            var targetObj = sourceSelection[0];
            var duplicatedObj = null;

            try {
                duplicatedObj = targetObj.duplicate();

                try {
                    duplicatedObj.filled = false;
                } catch (filledError) {
                }
                try {
                    duplicatedObj.stroked = false;
                } catch (strokedError) {
                }

                doc.selection = null;
                doc.selection = [duplicatedObj];
                app.redraw();
                app.executeMenuCommand('TrimMark v25');

                marks = doc.selection.slice(0);
            } catch (selectionTrimMarkError) {
                __logError('Failed to create trim marks for selected object', selectionTrimMarkError);
                throw selectionTrimMarkError;
            } finally {
                if (duplicatedObj) {
                    try {
                        duplicatedObj.remove();
                    } catch (removeDuplicatedSelectionError) {
                        __logError('Failed to remove duplicated selected object', removeDuplicatedSelectionError);
                    }
                }
            }

        } else if (targetMode === 'allArtboards') {

            marks = [];

            for (var abIndex = 0; abIndex < doc.artboards.length; abIndex++) {
                addTempItemFromArtboard(abIndex);

                var created = executeTrimMarkForCurrentTempItems();
                for (var c = 0; c < created.length; c++) {
                    marks.push(created[c]);
                }
            }

        } else {
            addTempItemFromArtboard(doc.artboards.getActiveArtboardIndex());
            marks = executeTrimMarkForCurrentTempItems();
        }

        if (marks.length === 0) {
            throw new Error(L('errorNoTarget'));
        }

        if (replaceLayer) {
            removeExistingTrimLayers();
        }

        /* トンボレイヤーを取得または作成 / Find or create the trim-mark layer */
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === 'トンボ') {
                trimLayer = doc.layers[i];
                break;
            }
        }
        if (!trimLayer) {
            trimLayer = doc.layers.add();
            trimLayer.name = 'トンボ';
        }

        /* トンボレイヤーを一時的にアンロック / Temporarily unlock the trim-mark layer */
        trimLayer.locked = false;

        /* 作成したトンボを移動 / Move created trim marks */
        for (var j = 0; j < marks.length; j++) {
            marks[j].move(trimLayer, ElementPlacement.PLACEATBEGINNING);
        }

        /* 選択解除してレイヤーをロック / Clear selection and lock the layer */
        doc.selection = null;
        trimLayer.locked = true;
    } catch (err) {
        /* エラー表示 / Show error */
        alert(err.message || err);
    } finally {
        try {
            doc.activeLayer = activeLayerBefore;
        } catch (restoreActiveLayerError) {
            __logError('Failed to restore active layer', restoreActiveLayerError);
        }

        for (var m = 0; m < tempItems.length; m++) {
            try {
                tempItems[m].remove();
            } catch (cleanupTempError) {
                __logError('Failed to remove temporary item in finally', cleanupTempError);
            }
        }
        cleanupTempLayer();

        try {
            app.preferences.setBooleanPreference('cropMarkStyle', cropMarkStyleBefore);
        } catch (prefRestoreError) {
            __logError('Failed to restore cropMarkStyle preference', prefRestoreError);
        }
        restoreAdjustedLayerStates(adjustedLayerStates);

        if (trimLayer) {
            /* トンボレイヤーをロック / Lock the trim-mark layer */
            try {
                trimLayer.locked = true;
            } catch (lockTrimLayerError) {
                __logError('Failed to lock trim layer', lockTrimLayerError);
            }
        }
    }
})();