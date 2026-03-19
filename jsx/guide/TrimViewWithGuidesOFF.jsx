#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * TrimViewWithGuidesOFF.jsx
 *
 * 概要:
 * 本スクリプトは、ガイド可視化用に作成されたプレビューレイヤー（内部識別子: __GUIDES_PREVIEW_TRIM_VIEW__）を削除し、
 * Trim View の表示状態を切り替えます。
 *
 * プレビューレイヤー内のロックされたオブジェクトは削除前に自動的に解除され、安全にレイヤーを削除します。
 *
 * 更新日: 2026-03-19
 */

var SCRIPT_VERSION = "v1.0";

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

    // 該当レイヤーを削除
    for (i = doc.layers.length - 1; i >= 0; i--) {
        var layer = doc.layers[i];

        if (layer && (layer.note === previewLayerNote || layer.name === previewLayerName)) {
            layer.locked = false;
            layer.visible = true;

            // 内部アイテムのロック解除
            var k;
            for (k = 0; k < layer.pageItems.length; k++) {
                var item = layer.pageItems[k];
                if (item && item.locked) {
                    try {
                        item.locked = false;
                    } catch (e) {}
                }
            }

            layer.remove();
        }
    }

    // トリム表示切り替え
    toggleTrimView();
})();
