#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * TrimViewWithGuidesON.jsx
 *
 * 概要:
 * ガイドを通常パスとして一時レイヤーに複製し、Trim View 切り替え時にもガイド位置を確認しやすくするスクリプトです。
 *
 * ガイドは「Guides Preview for Trim View」レイヤーに複製され、内部識別子として
 * __GUIDES_PREVIEW_TRIM_VIEW__ を設定します。
 * 既存の同名レイヤーまたは同一内部識別子を持つレイヤーがある場合には、それを削除してから新規に作成します。
 *
 * 更新日: 2026-03-19
 */

var SCRIPT_VERSION = "v1.0.1";
var GUIDE_STROKE_WIDTH = 1;

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var previewLayerName = "Guides Preview for Trim View";
    var previewLayerNote = "__GUIDES_PREVIEW_TRIM_VIEW__";
    var touchedLayers = [];
    var touchedItems = [];
    var guides = [];
    var i;

    // プレビュー用カラー（ドキュメントのカラーモードに応じてCMYK/RGB切替）
    function createPreviewColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var c = new CMYKColor();
            c.cyan = 0;
            c.magenta = 60;
            c.yellow = 0;
            c.black = 0;
            return c;
        } else {
            var r = new RGBColor();
            r.red = 227;
            r.green = 120;
            r.blue = 180;
            return r;
        }
    }

    // Trim View を切り替える
    function toggleTrimView() {
        app.executeMenuCommand('TrimView');
    }

    // スクリプトが生成したプレビューレイヤーかどうかを判定
    function isPreviewLayer(layer) {
        return layer && layer.note === previewLayerNote;
    }

    // 既存のプレビューレイヤーを検索
    function findExistingPreviewLayer() {
        var j;
        for (j = 0; j < doc.layers.length; j++) {
            var layer = doc.layers[j];
            if (isPreviewLayer(layer)) {
                return layer;
            }
        }
        return null;
    }

    // プレビューレイヤー削除前に内部アイテムのロックを解除
    function unlockItemsInLayer(layer) {
        var k;
        for (k = 0; k < layer.pageItems.length; k++) {
            if (layer.pageItems[k] && layer.pageItems[k].locked) {
                try {
                    layer.pageItems[k].locked = false;
                } catch (e) { }
            }
        }
    }

    var existingLayer = findExistingPreviewLayer();
    // 既存のプレビューレイヤーがある場合は削除してから続行
    if (existingLayer) {
        existingLayer.locked = false;
        existingLayer.visible = true;
        unlockItemsInLayer(existingLayer);
        existingLayer.remove();
    }

    // 新規レイヤー作成（ON）
    var previewLayer = doc.layers.add();
    previewLayer.name = previewLayerName;
    previewLayer.note = previewLayerNote;
    previewLayer.locked = false;
    previewLayer.visible = true;

    // 全レイヤーのロックを一時解除
    for (i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (layer.locked) {
            touchedLayers.push({ layer: layer, locked: true });
            layer.locked = false;
        }
    }

    // 全ページアイテムからガイドを収集
    for (i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];

        if (item) {
            if (item.locked) {
                try {
                    touchedItems.push({ item: item, locked: true });
                    item.locked = false;
                } catch (e) { }
            }

            if (item.guides === true) {
                guides.push(item);
            }
        }
    }

    // ガイドを複製して通常パス化＋線設定
    for (i = 0; i < guides.length; i++) {
        try {
            var dup = guides[i].duplicate(previewLayer, ElementPlacement.PLACEATBEGINNING);

            // ガイド解除
            dup.guides = false;

            // 線を有効化
            dup.stroked = true;
            dup.strokeWidth = GUIDE_STROKE_WIDTH;
            dup.strokeColor = createPreviewColor();

            // 塗りなし
            dup.filled = false;

            // レイヤー側で管理するため個別にはロックしない
            dup.locked = false;
        } catch (e) {
            // ガイド以外の特殊アイテムなど、複製できないものはスキップ
        }
    }

    // 元のアイテムのロック状態を復元
    for (i = 0; i < touchedItems.length; i++) {
        try {
            touchedItems[i].item.locked = touchedItems[i].locked;
        } catch (e) { }
    }

    // 元のレイヤーのロック状態を復元
    for (i = 0; i < touchedLayers.length; i++) {
        try {
            touchedLayers[i].layer.locked = touchedLayers[i].locked;
        } catch (e) { }
    }

    // プレビューレイヤー自体をロック
    previewLayer.locked = true;

    // トリム表示切り替え
    toggleTrimView();

})();
