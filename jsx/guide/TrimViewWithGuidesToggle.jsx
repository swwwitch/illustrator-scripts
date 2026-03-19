#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * TrimViewWithGuidesToggle.jsx
 *
 * 概要:
 * ガイドを通常パスとして一時レイヤー「__GUIDES_PREVIEW_TRIM_VIEW__」に複製し、Trim View でもガイド位置を確認しやすくするトグルスクリプトです。
 *
 * 同名レイヤーが存在しない場合はプレビューレイヤーを新規作成してガイドを複製し、最後に Trim View を切り替えます。
 * 同名レイヤーが存在する場合はその同名レイヤーをすべて削除して終了します。既存レイヤーの判定はレイヤー名一致のみを使用します。処理全体は finally でロック状態を復元し、ガイド複製時のみ対象ガイドを一時アンロックします。ガイド収集では非表示レイヤー上のガイドと非表示のガイド自体を対象外とし、ガイド本数・複製成功数・複製失敗数を集計します。ガイドが1本もない場合はプレビューレイヤーを作成せず終了します。
 *
 * 作成日：2026-03-19
 * 更新日: 2026-03-20
 */

var SCRIPT_VERSION = "v1.2";
var GUIDE_STROKE_WIDTH = 1;

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var previewLayerName = "__GUIDES_PREVIEW_TRIM_VIEW__";
    var touchedItems = [];
    var guides = [];
    var previewLayer = null;
    var guideCount = 0;
    var duplicateSuccessCount = 0;
    var duplicateFailureCount = 0;
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

    // 既存のプレビューレイヤーを検索（同名レイヤーをすべて対象にする）
    function findExistingPreviewLayers() {
        var j;
        var layer;
        var sameNameLayers = [];

        for (j = 0; j < doc.layers.length; j++) {
            layer = doc.layers[j];
            if (layer && layer.name === previewLayerName) {
                sameNameLayers.push(layer);
            }
        }

        return sameNameLayers;
    }

    // ガイドが属するすべての親レイヤーが表示中かどうかを判定
    function isGuideLayerVisible(item) {
        var parent = item ? item.parent : null;

        while (parent) {
            if (parent.typename === "Layer" && parent.visible === false) {
                return false;
            }
            parent = parent.parent;
        }

        return true;
    }

    // 既存の同名プレビューレイヤーがある場合は削除してから続行（複数一致はすべて削除）
    try {
        var existingLayers = findExistingPreviewLayers();
        var existingLayer;
        var cleared;

        if (existingLayers.length > 0) {
            for (i = 0; i < existingLayers.length; i++) {
                existingLayer = existingLayers[i];
                existingLayer.locked = false;

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
                    try {
                        existingLayer.remove();
                    } catch (e2) {
                        alert("空にしたプレビューレイヤーの削除に失敗しました。\n" +
                            "レイヤー名: " + existingLayer.name + "\n" +
                            "エラー: " + e2);
                        return;
                    }
                }
            }

            app.executeMenuCommand('TrimView');

            // 既存プレビューレイヤーがあった場合は OFF とみなして終了
            return;
        }

        // 全ページアイテムからガイドを収集
        for (i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            if (item && item.guides === true && item.hidden !== true && isGuideLayerVisible(item)) {
                guides.push(item);
            }
        }

        guideCount = guides.length;

        // ガイドが1本もない場合はプレビューレイヤーを作成しない
        if (guideCount === 0) {
            return;
        }

        // プレビューレイヤーを新規作成（ON）
        try {
            previewLayer = doc.layers.add();
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
        previewLayer.locked = false;

        // ガイドを必要なものだけ一時アンロックして複製
        for (i = 0; i < guides.length; i++) {
            var guide = guides[i];
            var wasLocked = false;
            var dup;

            if (!guide) {
                continue;
            }

            wasLocked = guide.locked;

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

        // プレビューレイヤー自体をロック
        previewLayer.locked = true;

        // Trim View を切り替え
        app.executeMenuCommand('TrimView');
    } finally {
        // 元のアイテムのロック状態を復元
        for (i = 0; i < touchedItems.length; i++) {
            try {
                touchedItems[i].item.locked = touchedItems[i].locked;
            } catch (e) { }
        }
    }

    // コンテナ自身と直下のページアイテムだけをアンロック
    function unlockItemsInContainer(container) {
        var k;
        var item;

        if (!container) {
            return;
        }

        if (container.locked) {
            container.locked = false;
        }

        if (!container.pageItems || container.pageItems.length === 0) {
            return;
        }

        for (k = 0; k < container.pageItems.length; k++) {
            item = container.pageItems[k];
            if (item && item.locked) {
                try {
                    item.locked = false;
                } catch (e) { }
            }
        }
    }

    // レイヤー削除に失敗した場合は中身を再帰的に空にしてから削除する
    function clearPreviewLayerContents(layer) {
        var k;
        var subLayer;
        var childCleared;

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

        if (layer.pageItems && layer.pageItems.length) {
            for (k = layer.pageItems.length - 1; k >= 0; k--) {
                try {
                    layer.pageItems[k].remove();
                } catch (e) {
                    return false;
                }
            }
        }

        return (layer.layers.length === 0 && layer.pageItems.length === 0);
    }

})();
