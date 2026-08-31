#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト（単純な長方形1点）、現在のアートボード、またはすべてのアートボードを対象に、トンボを作成します。
実行時のダイアログで、対象と「ガイドを残す」「日本式トンボ」のON/OFFを選べます。

詳細は README を参照してください。

### Overview

Creates trim marks for a selected object (a single simple rectangle), the current artboard, or every artboard.
A dialog picks the target and toggles "keep guides" and "Japanese-style trim marks".

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTrimMark";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMark.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMark.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n40e3e39cf9f2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function sanitizeLayerName(rawLayerName) {
        return rawLayerName.replace(new RegExp('[\\\\/:*?"<>|]', 'g'), '_');
    }

    function getOrCreateLayerByName(doc, layerName) {
        var i;
        for (i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
    }

    function buildTrimLayerNameForArtboard(artboard) {
        return "トンボ_" + sanitizeLayerName(artboard.name || "アートボード");
    }

    function unlockAndShowLayer(layer) {
        if (layer.locked) {
            layer.locked = false;
        }
        if (!layer.visible) {
            layer.visible = true;
        }
    }

    function createArtboardRectangle(trimLayer, artboard) {
        var artboardBounds = artboard.artboardRect;
        var rectItem = trimLayer.pathItems.rectangle(artboardBounds[1], artboardBounds[0], artboardBounds[2] - artboardBounds[0], artboardBounds[1] - artboardBounds[3]);
        rectItem.filled = false;
        rectItem.stroked = false;
        return rectItem;
    }

    function createTrimMarksFromSource(doc, trimLayer, sourceItem, keepGuides) {
        doc.activeLayer = trimLayer;
        doc.selection = null;
        doc.selection = [sourceItem];
        app.executeMenuCommand('TrimMark v25');

        if (keepGuides) {
            sourceItem.guides = true;
        } else {
            sourceItem.remove();
        }
    }

    function getUiLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUiLanguage();

    var LABELS = {
        dialogTitle: {
            ja: "トンボ作成",
            en: "Create Trim Marks"
        },
        panelTarget: {
            ja: "トンボの対象",
            en: "Trim Mark Target"
        },
        radioSelection: {
            ja: "選択オブジェクト",
            en: "Selected Object"
        },
        radioCurrentArtboard: {
            ja: "現在のアートボード",
            en: "Current Artboard"
        },
        radioAllArtboards: {
            ja: "すべてのアートボード",
            en: "All Artboards"
        },
        panelOptions: {
            ja: "オプション",
            en: "Options"
        },
        chkGuide: {
            ja: "ガイドを残す",
            en: "Keep Guides"
        },
        chkJapaneseTrim: {
            ja: "日本式トンボ",
            en: "Japanese-style Trim Marks"
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
        }
    };

    function getLabel(key) {
        return LABELS[key][uiLang];
    }

    // =========================================
    // Helper functions for rectangle validation
    // =========================================
    function isNearlyEqual(valueA, valueB) {
        return Math.abs(valueA - valueB) < 0.01;
    }

    function hasAnchorNear(anchorXs, anchorYs, cornerX, cornerY) {
        var i;
        for (i = 0; i < anchorXs.length; i++) {
            if (isNearlyEqual(anchorXs[i], cornerX) && isNearlyEqual(anchorYs[i], cornerY)) {
                return true;
            }
        }
        return false;
    }

    function isAxisAlignedRectangle(pathItem) {
        var pathPoints, anchorXs, anchorYs, i, left, right, top, bottom;

        if (pathItem.pathPoints.length !== 4 || !pathItem.closed) {
            return false;
        }

        pathPoints = pathItem.pathPoints;
        anchorXs = [];
        anchorYs = [];

        for (i = 0; i < 4; i++) {
            anchorXs.push(pathPoints[i].anchor[0]);
            anchorYs.push(pathPoints[i].anchor[1]);

            if (!isNearlyEqual(pathPoints[i].leftDirection[0], pathPoints[i].anchor[0]) ||
                !isNearlyEqual(pathPoints[i].leftDirection[1], pathPoints[i].anchor[1]) ||
                !isNearlyEqual(pathPoints[i].rightDirection[0], pathPoints[i].anchor[0]) ||
                !isNearlyEqual(pathPoints[i].rightDirection[1], pathPoints[i].anchor[1])) {
                return false;
            }
        }

        left = Math.min.apply(null, anchorXs);
        right = Math.max.apply(null, anchorXs);
        top = Math.max.apply(null, anchorYs);
        bottom = Math.min.apply(null, anchorYs);

        /* 幅または高さがゼロの退化パスを除外 / Reject degenerate paths with zero width or height */
        if (isNearlyEqual(left, right) || isNearlyEqual(top, bottom)) {
            return false;
        }

        for (i = 0; i < 4; i++) {
            if (!isNearlyEqual(anchorXs[i], left) && !isNearlyEqual(anchorXs[i], right)) {
                return false;
            }
            if (!isNearlyEqual(anchorYs[i], top) && !isNearlyEqual(anchorYs[i], bottom)) {
                return false;
            }
        }

        /* 4隅がすべて揃っているかを許容誤差込みで確認 / Check that all four corners are present, within tolerance */
        return hasAnchorNear(anchorXs, anchorYs, left, top) &&
            hasAnchorNear(anchorXs, anchorYs, right, top) &&
            hasAnchorNear(anchorXs, anchorYs, right, bottom) &&
            hasAnchorNear(anchorXs, anchorYs, left, bottom);
    }

    function isSimpleRectanglePathItem(item) {
        return item.typename === 'PathItem' && !item.guides && !item.clipping;
    }

    function isEligibleTrimSource(item) {
        return isSimpleRectanglePathItem(item) && isAxisAlignedRectangle(item);
    }

    // =========================================
    // ScriptUIダイアログ
    // =========================================
    function showOptionsDialog(doc) {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        var targetPanel = dialog.add('panel', undefined, getLabel('panelTarget'));
        var selectionRadio = targetPanel.add('radiobutton', undefined, getLabel('radioSelection'));
        var currentArtboardRadio = targetPanel.add('radiobutton', undefined, getLabel('radioCurrentArtboard'));
        var allArtboardsRadio = targetPanel.add('radiobutton', undefined, getLabel('radioAllArtboards'));
        var optionsPanel = dialog.add('panel', undefined, getLabel('panelOptions'));
        var keepGuidesCheckbox = optionsPanel.add('checkbox', undefined, getLabel('chkGuide'));
        var japaneseStyleCheckbox = optionsPanel.add('checkbox', undefined, getLabel('chkJapaneseTrim'));
        var btnRowGroup = dialog.add('group');
        var btnCancel = btnRowGroup.add('button', undefined, getLabel('btnCancel'), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel('btnOk'), { name: 'ok' });

        if (doc.selection.length === 1 && isEligibleTrimSource(doc.selection[0])) {
            selectionRadio.value = true;
        } else {
            currentArtboardRadio.value = true;
        }
        keepGuidesCheckbox.value = true;

        /* 現在の環境設定を初期値にする / Seed from the current preference */
        japaneseStyleCheckbox.value = app.preferences.getBooleanPreference('cropMarkStyle');

        targetPanel.orientation = 'column';
        targetPanel.alignChildren = 'left';
        targetPanel.alignment = 'fill';
        targetPanel.margins = [15, 20, 15, 10];
        optionsPanel.orientation = 'column';
        optionsPanel.alignChildren = 'left';
        optionsPanel.alignment = 'fill';
        optionsPanel.margins = [15, 20, 15, 10];
        btnRowGroup.alignment = 'right';
        btnCancel.preferredSize.width = 80;
        btnOK.preferredSize.width = 80;

        if (dialog.show() !== 1) {
            return null;
        }

        return {
            targetType: selectionRadio.value ? 'selection' : (allArtboardsRadio.value ? 'allArtboards' : 'artboard'),
            keepGuides: keepGuidesCheckbox.value,
            japaneseStyle: japaneseStyleCheckbox.value
        };
    }

    function getSelectedSimpleRectangle(doc) {
        var selectedItem;

        if (doc.selection.length !== 1) {
            alert(uiLang === 'ja'
                ? '「選択オブジェクト」を使うには、単純な長方形を1つだけ選択してください。'
                : 'To use "Selected Object", select exactly one simple rectangle.');
            return null;
        }

        selectedItem = doc.selection[0];

        if (!isSimpleRectanglePathItem(selectedItem)) {
            alert(uiLang === 'ja'
                ? '「選択オブジェクト」で使えるのは、単純な長方形だけです。'
                : 'Only a simple rectangle can be used for "Selected Object".');
            return null;
        }

        if (!isAxisAlignedRectangle(selectedItem)) {
            alert(uiLang === 'ja'
                ? '「選択オブジェクト」で使えるのは、各辺が水平・垂直で、4点が直交している単純な長方形だけです。'
                : 'Only a simple rectangle with horizontal/vertical edges and right-angle corners can be used for "Selected Object".');
            return null;
        }

        return selectedItem;
    }

    function duplicateItemToTrimLayer(sourceItem, trimLayer) {
        var duplicatedItem = sourceItem.duplicate(trimLayer, ElementPlacement.PLACEATBEGINNING);
        if ('filled' in duplicatedItem) {
            duplicatedItem.filled = false;
        }
        if ('stroked' in duplicatedItem) {
            duplicatedItem.stroked = false;
        }
        return duplicatedItem;
    }

    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alertNoDocument'));
            return;
        }

        var doc = app.activeDocument;
        var trimOptions = showOptionsDialog(doc);
        var selectedRectangle = null;
        var trimSourceItem = null;
        var trimLayer = null;
        var trimLayers = [];
        var trimLayerStates = [];
        var originalActiveLayer = doc.activeLayer;
        var originalJapaneseStyle = app.preferences.getBooleanPreference('cropMarkStyle');
        var didCreateTrimMarks = false;
        var i;

        if (!trimOptions) {
            return;
        }

        /* レイヤーや環境設定を変更する前に選択オブジェクトを検証 / Validate the selection before touching layers or preferences */
        if (trimOptions.targetType === 'selection') {
            selectedRectangle = getSelectedSimpleRectangle(doc);
            if (!selectedRectangle) {
                return;
            }
        }

        if (trimOptions.targetType !== 'allArtboards') {
            /* 「トンボ」レイヤーを取得（なければ作成） / Get the "Trim" layer, or create it if missing */
            trimLayer = getOrCreateLayerByName(doc, "トンボ");
            trimLayers.push(trimLayer);
            trimLayerStates.push({ locked: trimLayer.locked, visible: trimLayer.visible });
            unlockAndShowLayer(trimLayer);
        }

        try {

            /* 日本式トンボの設定を切り替える / Set Japanese-style trim marks */
            app.preferences.setBooleanPreference('cropMarkStyle', trimOptions.japaneseStyle ? 1 : 0);

            if (trimOptions.targetType === 'selection') {
                /* 選択オブジェクトを複製してトンボ用に準備 / Duplicate the selected object for trim marks */
                trimSourceItem = duplicateItemToTrimLayer(selectedRectangle, trimLayer);
                createTrimMarksFromSource(doc, trimLayer, trimSourceItem, trimOptions.keepGuides);
            } else if (trimOptions.targetType === 'allArtboards') {
                for (i = 0; i < doc.artboards.length; i++) {
                    trimLayer = getOrCreateLayerByName(doc, buildTrimLayerNameForArtboard(doc.artboards[i]));
                    trimLayers.push(trimLayer);
                    trimLayerStates.push({ locked: trimLayer.locked, visible: trimLayer.visible });
                    unlockAndShowLayer(trimLayer);
                    trimSourceItem = createArtboardRectangle(trimLayer, doc.artboards[i]);
                    createTrimMarksFromSource(doc, trimLayer, trimSourceItem, trimOptions.keepGuides);
                }
            } else {
                /* アクティブなアートボードを取得 / Get the active artboard */
                var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];

                /* アートボード矩形を作成 / Create the artboard rectangle */
                trimSourceItem = createArtboardRectangle(trimLayer, activeArtboard);
                createTrimMarksFromSource(doc, trimLayer, trimSourceItem, trimOptions.keepGuides);
            }

            didCreateTrimMarks = true;

        } finally {
            /* トンボを作成できたときだけ選択を解除 / Deselect only when trim marks were actually created */
            if (didCreateTrimMarks) {
                doc.selection = null;
            }

            /* 環境設定を元に戻す / Restore the preference */
            app.preferences.setBooleanPreference('cropMarkStyle', originalJapaneseStyle ? 1 : 0);

            /* ロックや非表示を戻す前にアクティブレイヤーを復帰 / Restore the active layer before re-locking or re-hiding */
            if (originalActiveLayer.visible) {
                doc.activeLayer = originalActiveLayer;
            }

            for (i = 0; i < trimLayers.length; i++) {
                trimLayers[i].visible = trimLayerStates[i].visible;
                trimLayers[i].locked = trimLayerStates[i].locked;
            }
        }
    }

    main();

})();
