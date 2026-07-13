#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
TrimViewWithGuidesOFF.jsx

概要:
本スクリプトは、ガイド可視化用に作成されたプレビューレイヤー（内部識別子: __GUIDES_PREVIEW_TRIM_VIEW__）を削除し、
Trim View の表示状態を切り替えます。

既存レイヤーの削除判定は内部識別子（note一致）を最優先し、note一致が複数ある場合はそのすべてを削除対象にします。
note一致がない場合のみ同名レイヤーを削除対象として扱います。

プレビューレイヤー削除前には、内部アイテムとサブレイヤーのロック・非表示状態を再帰的に解除します。
レイヤー削除に失敗した場合は中身を空にし、処理全体の後始末は finally で実行します。

作成日: 2026-03-19
更新日: 2026-03-20
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.1";

// =========================================
// ローカライズ / Localization
// =========================================

var lang = ($.locale && /^ja/i.test($.locale)) ? "ja" : "en";

function L(labelText) {
    return (labelText && labelText[lang]) || (labelText && labelText.ja) || "";
}

var LABELS = {
    err_remove_preview: {
        ja: "プレビューレイヤーの削除に失敗し、中身のクリアにも失敗しました。",
        en: "Failed to remove the preview layer; clearing its contents also failed."
    },
    field_layer_name: { ja: "レイヤー名: ", en: "Layer: " },
    field_error: { ja: "エラー: ", en: "Error: " }
};

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var previewLayerNote = "__GUIDES_PREVIEW_TRIM_VIEW__";
    var previewLayerName = "Guides Preview for Trim View";

    var i;

    /* Trim View 表示の ON/OFF を切り替え / Toggle Trim View on/off */
    function toggleTrimView() {
        app.executeMenuCommand('TrimView');
    }

    /* スクリプトが生成したプレビューレイヤーかどうかを判定 / Check if a layer was created as a preview layer by this script */
    function isPreviewLayer(layer) {
        return layer && layer.note === previewLayerNote;
    }

    /* 既存プレビューレイヤーを検索（note 一致を最優先、なければ同名を対象） / Find existing preview layers (note match preferred, otherwise same-name) */
    function findExistingPreviewLayers() {
        var j;
        var layer;
        var sameNameLayers = [];
        var noteMatchedLayers = [];

        for (j = 0; j < doc.layers.length; j++) {
            layer = doc.layers[j];
            if (isPreviewLayer(layer)) {
                noteMatchedLayers.push(layer);
            } else if (layer && layer.name === previewLayerName) {
                sameNameLayers.push(layer);
            }
        }

        if (noteMatchedLayers.length > 0) {
            return noteMatchedLayers;
        }

        return sameNameLayers;
    }

    /* コンテナ内のアイテムとサブレイヤーのロック・非表示を再帰的に解除 / Recursively unlock/unhide items and sublayers inside a container */
    function unlockItemsInContainer(container) {
        var k;
        var pageItem;
        var subLayer;

        if (!container) {
            return;
        }

        if (container.locked) {
            try {
                container.locked = false;
            } catch (e) { }
        }

        if (container.visible === false) {
            try {
                container.visible = true;
            } catch (e) { }
        }

        if (container.layers && container.layers.length) {
            for (k = 0; k < container.layers.length; k++) {
                subLayer = container.layers[k];
                unlockItemsInContainer(subLayer);
            }
        }

        if (container.pageItems && container.pageItems.length) {
            for (k = 0; k < container.pageItems.length; k++) {
                pageItem = container.pageItems[k];
                if (!pageItem) {
                    continue;
                }

                if (pageItem.locked) {
                    try {
                        pageItem.locked = false;
                    } catch (e) { }
                }

                if (pageItem.hidden) {
                    try {
                        pageItem.hidden = false;
                    } catch (e) { }
                }
            }
        }
    }

    /* コンテナ内のページアイテムを末尾から削除 / Remove pageItems in a container from the tail */
    function removePageItemsInContainer(container) {
        var k;

        if (!container || !container.pageItems) {
            return true;
        }

        for (k = container.pageItems.length - 1; k >= 0; k--) {
            try {
                container.pageItems[k].remove();
            } catch (e) {
                return false;
            }
        }

        return true;
    }

    /* レイヤー削除失敗時の後始末として中身を空にする / Empty a layer's contents as fallback when layer removal fails */
    function clearPreviewLayerContents(layer) {
        var k;
        var subLayer;

        if (!layer) {
            return false;
        }

        unlockItemsInContainer(layer);

        if (layer.layers && layer.layers.length) {
            for (k = layer.layers.length - 1; k >= 0; k--) {
                subLayer = layer.layers[k];
                if (!clearPreviewLayerContents(subLayer)) {
                    return false;
                }
                try {
                    subLayer.remove();
                } catch (e) {
                    return false;
                }
            }
        }

        if (!removePageItemsInContainer(layer)) {
            return false;
        }

        return true;
    }

    var existingLayers = findExistingPreviewLayers();
    var existingLayer;

    // 既存のプレビューレイヤーがある場合は削除（note一致を最優先し、複数一致はすべて削除）
    if (existingLayers.length > 0) {
        for (i = 0; i < existingLayers.length; i++) {
            existingLayer = existingLayers[i];
            existingLayer.locked = false;
            existingLayer.visible = true;
            unlockItemsInContainer(existingLayer);

            try {
                existingLayer.remove();
            } catch (e) {
                if (!clearPreviewLayerContents(existingLayer)) {
                    alert(L(LABELS.err_remove_preview) + "\n" +
                        L(LABELS.field_layer_name) + existingLayer.name + "\n" +
                        L(LABELS.field_error) + e);
                    return;
                }
            }
        }
    }

    // トリム表示切り替え
    toggleTrimView();
})();
