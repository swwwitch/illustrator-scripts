#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * TrimViewWithGuidesOFF.jsx
 *
 * 概要:
 * 本スクリプトは、ガイド可視化用に作成されたプレビューレイヤー（内部識別子: __GUIDES_PREVIEW_TRIM_VIEW__）を削除し、
 * Trim View の表示状態を切り替えます。
 *
 * 既存レイヤーの削除判定は内部識別子（note一致）を最優先し、note一致が複数ある場合はそのすべてを削除対象にします。
 * note一致がない場合のみ同名レイヤーを削除対象として扱います。
 *
 * プレビューレイヤー削除前には、内部アイテムとサブレイヤーのロック・非表示状態を再帰的に解除します。
 * レイヤー削除に失敗した場合は中身を空にし、処理全体の後始末は finally で実行します。
 * 
 * 作成日: 2026-03-19
 * 更新日: 2026-03-20
 */

var SCRIPT_VERSION = "v1.1";

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var previewLayerNote = "__GUIDES_PREVIEW_TRIM_VIEW__";
    var previewLayerName = "Guides Preview for Trim View";

    var i;

    // Trim View を切り替える
    function toggleTrimView() {
        app.executeMenuCommand('TrimView');
    }

    function isPreviewLayer(layer) {
        return layer && layer.note === previewLayerNote;
    }

    // 既存のプレビューレイヤーを検索（note一致を最優先し、なければ同名を対象にする）
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

    // プレビューレイヤー削除前に内部アイテムとサブレイヤーのロックを再帰的に解除
    function unlockItemsInContainer(container) {
        var k;
        var item;
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
                item = container.pageItems[k];
                if (!item) {
                    continue;
                }

                if (item.locked) {
                    try {
                        item.locked = false;
                    } catch (e) { }
                }

                if (item.hidden) {
                    try {
                        item.hidden = false;
                    } catch (e) { }
                }
            }
        }
    }

    // コンテナ内のページアイテムを末尾から削除
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

    // レイヤー削除に失敗した場合は中身を空にする
    function clearPreviewLayerContents(layer) {
        var k;
        var subLayer;
        var childCleared;
        var removed;

        if (!layer) {
            return false;
        }

        unlockItemsInContainer(layer);

        if (layer.layers && layer.layers.length) {
            for (k = layer.layers.length - 1; k >= 0; k--) {
                subLayer = layer.layers[k];
                childCleared = clearPreviewLayerContents(subLayer);
                if (!childCleared) {
                    return false;
                }
                try {
                    subLayer.remove();
                } catch (e) {
                    return false;
                }
            }
        }

        removed = removePageItemsInContainer(layer);
        if (!removed) {
            return false;
        }

        return (layer.layers.length === 0 && layer.pageItems.length === 0);
    }

    var touchedLayers = [];
    var existingLayers;
    var existingLayer;
    var cleared;

    try {
        // 全レイヤーのロックを一時解除
        for (i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (layer.locked) {
                touchedLayers.push({ layer: layer, locked: true });
                layer.locked = false;
            }
        }

        existingLayers = findExistingPreviewLayers();

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
                    cleared = clearPreviewLayerContents(existingLayer);
                    if (!cleared) {
                        alert("プレビューレイヤーの削除に失敗し、中身のクリアにも失敗しました。\n" +
                            "レイヤー名: " + existingLayer.name + "\n" +
                            "エラー: " + e);
                        return;
                    }
                }
            }
        }

        // トリム表示切り替え
        toggleTrimView();
    } finally {
        for (i = 0; i < touchedLayers.length; i++) {
            try {
                touchedLayers[i].layer.locked = touchedLayers[i].locked;
            } catch (e) { }
        }
    }
})();
