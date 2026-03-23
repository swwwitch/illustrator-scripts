#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * TrimViewWithGuidesON.jsx
 *
 * 概要:
 * ガイドを通常パスとして専用レイヤーに複製し、Trim View 切り替え時にもガイド位置を確認しやすくするスクリプトです。
 *
 * ガイドは「Guides Preview for Trim View」レイヤーに複製され、内部識別子として
 * __GUIDES_PREVIEW_TRIM_VIEW__ を設定します。既存の同レイヤーがある場合は事前に削除し、削除に失敗した場合のみ
 * 中身を空にして再利用します。ガイドの複製は通常レイヤーの状態で行い、複製完了後にテンプレートレイヤー化します。
 * note一致を最優先で削除対象にし、note一致がない場合のみ同名レイヤーを対象にします。note一致が複数ある場合はそのすべてを削除対象にします。
 * 処理全体は finally でロック状態を復元し、ガイド複製時のみ対象ガイドを一時アンロックします。非表示レイヤー配下のガイドと非表示のガイドは対象外です。
 * ガイドが1本もない場合はプレビューレイヤーを作成せず終了します。
 *
 * 作成日：2026-03-19
 * 更新日: 2026-03-23
 */

var SCRIPT_VERSION = "v1.2";
var GUIDE_STROKE_WIDTH = 1;

var TEMPLATE_ACTION_NAME = "template";
var TEMPLATE_ACTION_SET_NAME = "layer";
var TEMPLATE_ACTION_FILE_NAME = "TrimViewWithGuidesTemplateAction.aia";

function loadAndRunAction(actionString, actionName, actionSetName) {
    var f = new File(Folder.temp + '/' + TEMPLATE_ACTION_FILE_NAME);

    try {
        f.open('w');
        f.write(actionString);
        f.close();

        app.loadAction(f);
        app.doScript(actionName, actionSetName, false);
        app.unloadAction(actionSetName, '');
    } finally {
        try {
            if (f.exists) {
                f.remove();
            }
        } catch (_) { }
    }
}

function getTemplateLayerActionString() {
    return '/version 3/name [ 5 6c61796572 ]/isOpen 1/actionCount 1/action-1 { /name [ 8 74656d706c617465 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_Layer) /localizedName [ 9 e8a1a8e7a4ba203a20 ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 9 /parameter-1 { /key 1836411236 /showInPalette 4294967295 /type (integer) /value 4 } /parameter-2 { /key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 36 e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383a7e383b3 ] } /parameter-3 { /key 1953329260 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-4 { /key 1936224119 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-5 { /key 1819239275 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-6 { /key 1886549623 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-7 { /key 1886547572 /showInPalette 4294967295 /type (boolean) /value 0 } /parameter-8 { /key 1684630830 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-9 { /key 1885564532 /showInPalette 4294967295 /type (unit real) /value 50.0 /unit 592474723 } }}';
}

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
    var previewLayer = null;
    var guideCount = 0;
    var duplicateSuccessCount = 0;
    var duplicateFailureCount = 0;
    var i;
    var docSelection = null;

    // 指定レイヤーをテンプレートレイヤーに変換
    function makeLayerTemplate(targetLayer) {
        if (!targetLayer) {
            return;
        }

        docSelection = doc.selection;
        doc.selection = null;
        doc.activeLayer = targetLayer;
        loadAndRunAction(getTemplateLayerActionString(), TEMPLATE_ACTION_NAME, TEMPLATE_ACTION_SET_NAME);
        targetLayer.name = previewLayerName;
        targetLayer.note = previewLayerNote;
    }

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

    // 親 Layer を上にたどり、どこか1つでも非表示なら true
    function isInHiddenLayer(item) {
        var parent = item ? item.parent : null;

        while (parent) {
            if (parent.typename === "Layer" && parent.visible === false) {
                return true;
            }
            parent = parent.parent;
        }

        return false;
    }

    // 既存のプレビューレイヤーを検索（note一致を最優先し、なければ同名を対象にする）
    function findExistingPreviewLayers() {
        var j;
        var layer;
        var sameNameLayer = null;
        var noteMatchedLayers = [];

        for (j = 0; j < doc.layers.length; j++) {
            layer = doc.layers[j];
            if (isPreviewLayer(layer)) {
                noteMatchedLayers.push(layer);
            } else if (!sameNameLayer && layer && layer.name === previewLayerName) {
                sameNameLayer = layer;
            }
        }

        if (noteMatchedLayers.length > 0) {
            return noteMatchedLayers;
        }

        return sameNameLayer ? [sameNameLayer] : [];
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

    // レイヤー削除に失敗した場合は中身を空にして再利用する
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

    try {
        // 全レイヤーのロックを一時解除
        for (i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (layer.locked) {
                touchedLayers.push({ layer: layer, locked: true });
                layer.locked = false;
            }
        }

        var existingLayers = findExistingPreviewLayers();
        var existingLayer;
        var reusablePreviewLayer = null;
        var cleared;

        // 既存のプレビューレイヤーがある場合は削除してから続行（note一致を最優先し、複数一致はすべて削除）
        if (existingLayers.length > 0) {
            for (i = 0; i < existingLayers.length; i++) {
                existingLayer = existingLayers[i];
                existingLayer.locked = false;
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
                    if (!reusablePreviewLayer) {
                        reusablePreviewLayer = existingLayer;
                    }
                }
            }
        }

        // 全ページアイテムからガイドを収集（非表示レイヤー配下・非表示ガイドは除外）
        for (i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            if (!item || item.guides !== true) {
                continue;
            }
            if (item.hidden) {
                continue;
            }
            if (isInHiddenLayer(item)) {
                continue;
            }
            guides.push(item);
        }

        guideCount = guides.length;
        duplicateSuccessCount = 0;
        duplicateFailureCount = 0;

        // ガイドが1本もない場合はプレビューレイヤーを作成しない
        if (guideCount === 0) {
            return;
        }

        // プレビューレイヤーを作成または再利用（ON）
        try {
            previewLayer = reusablePreviewLayer || doc.layers.add();
        } catch (e) {
            alert("プレビューレイヤーの作成に失敗しました。\n" +
                "エラー: " + e);
            return;
        }

        if (!previewLayer || previewLayer.typename !== "Layer") {
            alert("プレビューレイヤーの作成に失敗しました。\n" +
                "Layer オブジェクトを取得できませんでした。");
            return;
        }

        previewLayer.name = previewLayerName;
        previewLayer.note = previewLayerNote;
        previewLayer.locked = false;

        // ガイドを必要なものだけ一時アンロックして複製
        for (i = 0; i < guides.length; i++) {
            var guide = guides[i];
            var wasLocked = false;
            var dup;

            if (!guide) {
                continue;
            }

            try {
                wasLocked = guide.locked;
            } catch (e) {
                wasLocked = false;
            }

            if (wasLocked) {
                try {
                    touchedItems.push({ item: guide, locked: true });
                    guide.locked = false;
                } catch (e) {
                    continue;
                }
            }

            try {
                dup = guide.duplicate(previewLayer, ElementPlacement.PLACEATBEGINNING);

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
                duplicateSuccessCount++;
            } catch (e) {
                duplicateFailureCount++;
            }
        }

        if (guideCount > 0 && duplicateSuccessCount === 0) {
            alert("ガイドのプレビューを作成できませんでした。\n" +
                "ガイド本数: " + guideCount + "\n" +
                "複製成功数: " + duplicateSuccessCount + "\n" +
                "複製失敗数: " + duplicateFailureCount);
        }

        // 複製完了後にテンプレートレイヤー化
        makeLayerTemplate(previewLayer);

        // // プレビューレイヤー自体をロック
        // previewLayer.locked = true;

        // トリム表示切り替え
        toggleTrimView();
    } finally {
        try {
            if (docSelection !== null) {
                doc.selection = docSelection;
            }
        } catch (e) { }

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
    }

})();
