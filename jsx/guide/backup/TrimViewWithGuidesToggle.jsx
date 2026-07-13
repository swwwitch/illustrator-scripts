#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
TrimViewWithGuidesToggle.jsx

概要:
「Guides Preview for Trim View」レイヤーの有無に応じて、ガイドプレビュー表示をトグルするスクリプトです。

既存の同レイヤー（__GUIDES_PREVIEW_TRIM_VIEW__）がある場合は、そのレイヤーを削除して Trim View を切り替えて終了します。
既存レイヤーがない場合は、表示中のガイドだけを通常パスとして専用レイヤーに複製し、複製完了後にテンプレートレイヤー化します。
テンプレートレイヤー化は、プレビューレイヤーをアクティブにした状態で動的アクションを実行し、変換対象を「Guides Preview for Trim View」レイヤーのみに限定します。

既存レイヤーの検出は note 一致を最優先とし、なければ同名レイヤーを対象にします。複数一致した場合はすべて削除対象です。
既存プレビューレイヤーの削除に失敗した場合は、中身のクリアを試み、継続不能なら中止します。
非表示レイヤー配下のガイドおよび非表示ガイドは対象外です。
レイヤーの一括アンロックは行わず、既存プレビューレイヤー削除時とガイド複製時に必要な対象だけを個別にアンロックします。
ガイド複製中に1件でも失敗した場合は、途中まで作成したプレビューの後始末を試みた上で中止し、後始末失敗も通知します。

紹介記事（note）
https://note.com/dtp_tranist/n/n338bdcc94636

作成日：2026-03-19
更新日: 2026-03-25
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.3.3";

// =========================================
// ユーザー設定 / User settings
// =========================================

var GUIDE_STROKE_WIDTH = 1;

var TEMPLATE_ACTION_NAME = "template";
var TEMPLATE_ACTION_SET_NAME = "layer";
var TEMPLATE_ACTION_FILE_NAME = "TrimViewWithGuidesTemplateAction.aia";

// =========================================
// ローカライズ / Localization
// =========================================

var lang = ($.locale && /^ja/i.test($.locale)) ? "ja" : "en";

function L(labelText) {
    return (labelText && labelText[lang]) || (labelText && labelText.ja) || "";
}

var LABELS = {
    err_unlock_existing: {
        ja: "既存プレビューレイヤーのロック解除に失敗しました。",
        en: "Failed to unlock the existing preview layer."
    },
    err_remove_preview: {
        ja: "プレビューレイヤーの削除に失敗し、中身のクリアにも失敗しました。",
        en: "Failed to remove the preview layer; clearing its contents also failed."
    },
    err_create_preview: {
        ja: "プレビューレイヤーの作成に失敗しました。",
        en: "Failed to create the preview layer."
    },
    err_create_preview_invalid: {
        ja: "Layer オブジェクトを取得できませんでした。",
        en: "Could not obtain a Layer object."
    },
    err_duplicate_aborted: {
        ja: "ガイドの複製中にエラーが発生したため、プレビューの作成を中止しました。",
        en: "An error occurred while duplicating guides; preview creation has been aborted."
    },
    err_preview_failed: {
        ja: "ガイドのプレビューを作成できませんでした。",
        en: "Could not create the guide preview."
    },
    err_cleanup_failed: {
        ja: "プレビューレイヤーの後始末にも失敗しました。",
        en: "Cleanup of the preview layer also failed."
    },
    field_layer_name: { ja: "レイヤー名: ", en: "Layer: " },
    field_error: { ja: "エラー: ", en: "Error: " },
    field_guide_count: { ja: "ガイド本数: ", en: "Guide count: " },
    field_success_count: { ja: "複製成功数: ", en: "Success count: " },
    field_failure_count: { ja: "複製失敗数: ", en: "Failure count: " }
};

// =========================================
// 一時アクション生成 / Temporary action generation
// =========================================

/* テンプレートレイヤー化アクションのソースを生成 / Build action source for converting a layer to template */
function buildTemplateLayerActionSource(setName, actionName) {
    return ''
        + '/version 3\n'
        + buildActionNameLine(setName)
        + '/isOpen 1\n'
        + '/actionCount 1\n'
        + '/action-1 {\n'
        + buildActionNameLine(actionName)
        + ' /keyIndex 0\n'
        + ' /colorIndex 0\n'
        + ' /isOpen 1\n'
        + ' /eventCount 1\n'
        + ' /event-1 {\n'
        + ' /useRulersIn1stQuadrant 0\n'
        + ' /internalName (ai_plugin_Layer)\n'
        + ' /localizedName [ 9 e8a1a8e7a4ba203a20 ]\n'
        + ' /isOpen 1\n'
        + ' /isOn 1\n'
        + ' /hasDialog 1\n'
        + ' /showDialog 0\n'
        + ' /parameterCount 9\n'
        + ' /parameter-1 { /key 1836411236 /showInPalette 4294967295 /type (integer) /value 4 }\n'
        + ' /parameter-2 { /key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 36 e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383a7e383b3 ] }\n'
        + ' /parameter-3 { /key 1953329260 /showInPalette 4294967295 /type (boolean) /value 1 }\n'
        + ' /parameter-4 { /key 1936224119 /showInPalette 4294967295 /type (boolean) /value 1 }\n'
        + ' /parameter-5 { /key 1819239275 /showInPalette 4294967295 /type (boolean) /value 1 }\n'
        + ' /parameter-6 { /key 1886549623 /showInPalette 4294967295 /type (boolean) /value 1 }\n'
        + ' /parameter-7 { /key 1886547572 /showInPalette 4294967295 /type (boolean) /value 0 }\n'
        + ' /parameter-8 { /key 1684630830 /showInPalette 4294967295 /type (boolean) /value 1 }\n'
        + ' /parameter-9 { /key 1885564532 /showInPalette 4294967295 /type (unit real) /value 50.0 /unit 592474723 }\n'
        + ' }\n'
        + '}\n';
}

/* アクション名行を生成（長さ＋ヘキサ） / Build the action /name line (length + hex) */
function buildActionNameLine(actionName) {
    return '/name [ ' + actionName.length + ' ' + stringToHex(actionName) + ' ]\n';
}

/* 文字列をヘキサ文字列に変換 / Convert a string to a hex string */
function stringToHex(sourceText) {
    var hexText = "";
    for (var i = 0; i < sourceText.length; i++) {
        var hexValue = sourceText.charCodeAt(i).toString(16);
        if (hexValue.length < 2) hexValue = "0" + hexValue;
        hexText += hexValue;
    }
    return hexText;
}

// =========================================
// 一時アクション実行 / Temporary action playback
// =========================================

/* 一時アクションを生成・実行・後始末 / Create, run, and clean up a temporary action */
function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
    var actionFile = new File(actionFilePath);
    var isActionLoaded = false;
    var isActionFileOpen = false;

    try { app.unloadAction(setName, ""); } catch (e) { }

    try {
        if (!actionFile.open("w")) {
            throw new Error("Failed to open temporary action file.");
        }
        isActionFileOpen = true;

        actionFile.write(actionSource);
        actionFile.close();
        isActionFileOpen = false;

        app.loadAction(actionFile);
        isActionLoaded = true;

        app.doScript(actionName, setName, false);

    } finally {
        if (isActionFileOpen) {
            try { actionFile.close(); } catch (e) { }
        }

        if (actionFile.exists) {
            try { actionFile.remove(); } catch (e) { }
        }

        if (isActionLoaded) {
            try { app.unloadAction(setName, ""); } catch (e) { }
        }
    }
}

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var previewLayerName = "Guides Preview for Trim View";
    var previewLayerNote = "__GUIDES_PREVIEW_TRIM_VIEW__";
    var guidesToRelock = [];
    var guides = [];
    var previewLayer = null;
    var guideCount = 0;
    var duplicateSuccessCount = 0;
    var duplicateFailureCount = 0;
    var i;
    var docSelection = null;
    var originalActiveLayer = null;
    var previewCleanupFailed = false;

    /* 指定レイヤーをテンプレートレイヤーに変換 / Convert the given layer to a template layer */
    function makeLayerTemplate(targetLayer) {
        if (!targetLayer) {
            return;
        }

        docSelection = doc.selection;
        originalActiveLayer = doc.activeLayer;
        doc.selection = null;
        doc.activeLayer = targetLayer;

        var actionSource = buildTemplateLayerActionSource(TEMPLATE_ACTION_SET_NAME, TEMPLATE_ACTION_NAME);
        var actionFilePath = Folder.temp + '/' + TEMPLATE_ACTION_FILE_NAME;
        playTemporaryAction(actionSource, TEMPLATE_ACTION_SET_NAME, TEMPLATE_ACTION_NAME, actionFilePath);

        targetLayer.name = previewLayerName;
        targetLayer.note = previewLayerNote;
    }

    /* ガイドをプレビューレイヤーに複製してパスとして整形 / Duplicate a guide to the preview layer and style it as a stroked path */
    function duplicateGuideToPreview(guide, previewLayer) {
        var previewPath = guide.duplicate(previewLayer, ElementPlacement.PLACEATBEGINNING);
        previewPath.guides = false;
        previewPath.stroked = true;
        previewPath.strokeWidth = GUIDE_STROKE_WIDTH;
        previewPath.strokeColor = createPreviewColor();
        previewPath.filled = false;
        previewPath.locked = false;
        return previewPath;
    }

    /* プレビュー用カラーを生成（CMYK/RGB をドキュメントモードで切替） / Build preview color (CMYK or RGB depending on document mode) */
    function createPreviewColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 60;
            cmykColor.yellow = 0;
            cmykColor.black = 0;
            return cmykColor;
        } else {
            var rgbColor = new RGBColor();
            rgbColor.red = 227;
            rgbColor.green = 120;
            rgbColor.blue = 180;
            return rgbColor;
        }
    }

    /* Trim View 表示の ON/OFF を切り替え / Toggle Trim View on/off */
    function toggleTrimView() {
        app.executeMenuCommand('TrimView');
    }

    /* 既存プレビューレイヤーを順に削除（失敗時は中身クリアでフォールバック） / Remove existing preview layers one by one (fallback to content clearing on failure) */
    function removeExistingPreviewLayers(layers) {
        for (var idx = 0; idx < layers.length; idx++) {
            var existingLayer = layers[idx];
            try {
                existingLayer.locked = false;
            } catch (e) {
                alert(L(LABELS.err_unlock_existing) + "\n" +
                    L(LABELS.field_layer_name) + existingLayer.name + "\n" +
                    L(LABELS.field_error) + e);
                return false;
            }

            unlockItemsInContainer(existingLayer);

            try {
                existingLayer.remove();
            } catch (e) {
                if (!clearPreviewLayerContents(existingLayer)) {
                    alert(L(LABELS.err_remove_preview) + "\n" +
                        L(LABELS.field_layer_name) + existingLayer.name + "\n" +
                        L(LABELS.field_error) + e);
                    return false;
                }
            }
        }
        return true;
    }

    /* 表示中ガイドを収集（非表示・非表示レイヤー配下を除外） / Collect visible guides (excluding hidden ones and those under hidden layers) */
    function collectVisibleGuides() {
        var collected = [];
        for (var idx = 0; idx < doc.pageItems.length; idx++) {
            var pageItem = doc.pageItems[idx];
            if (!pageItem || pageItem.guides !== true) continue;
            if (pageItem.hidden) continue;
            if (isInHiddenLayer(pageItem)) continue;
            collected.push(pageItem);
        }
        return collected;
    }

    /* プレビューレイヤーを新規作成して名前・note・ロック解除を設定 / Create a new preview layer with name/note/unlock set */
    function createPreviewLayer() {
        var layer;
        try {
            layer = doc.layers.add();
        } catch (e) {
            alert(L(LABELS.err_create_preview) + "\n" +
                L(LABELS.field_error) + e);
            return null;
        }

        if (!layer || layer.typename !== "Layer") {
            alert(L(LABELS.err_create_preview) + "\n" +
                L(LABELS.err_create_preview_invalid));
            return null;
        }

        layer.name = previewLayerName;
        layer.note = previewLayerNote;
        layer.locked = false;
        return layer;
    }

    /* ガイドを順に複製。1 件でも失敗したら break / Duplicate guides in order; break on first failure */
    function duplicateGuidesToPreview() {
        for (var idx = 0; idx < guides.length; idx++) {
            var guide = guides[idx];
            if (!guide) continue;

            var wasLocked = false;
            try { wasLocked = guide.locked; } catch (e) { }

            if (wasLocked) {
                try {
                    guidesToRelock.push(guide);
                    guide.locked = false;
                } catch (e) {
                    continue;
                }
            }

            try {
                duplicateGuideToPreview(guide, previewLayer);
                duplicateSuccessCount++;
            } catch (e) {
                duplicateFailureCount++;
                break;
            }
        }
    }

    /* スクリプトが生成したプレビューレイヤーかどうかを判定 / Check if a layer was created as a preview layer by this script */
    function isPreviewLayer(layer) {
        return layer && layer.note === previewLayerNote;
    }

    /* 親 Layer を上にたどり、どれか 1 つでも非表示なら true / Walk up parents and return true if any ancestor Layer is hidden */
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

    /* 既存プレビューレイヤーを検索（note 一致を最優先、なければ同名を対象） / Find existing preview layers (note match preferred, otherwise same-name) */
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

    /* プレビューレイヤーを「中身クリア＋削除」で完全に片付ける / Fully clean up a preview layer (clear contents then remove) */
    function cleanupPartialPreview(layer) {
        if (!clearPreviewLayerContents(layer)) {
            return false;
        }
        try {
            layer.remove();
        } catch (e) {
            return false;
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

    try {
        var existingLayers = findExistingPreviewLayers();

        // 既存プレビューがあれば削除して Trim View をトグルして終了
        if (existingLayers.length > 0) {
            if (!removeExistingPreviewLayers(existingLayers)) {
                return;
            }
            toggleTrimView();
            return;
        }

        guides = collectVisibleGuides();
        guideCount = guides.length;

        // ガイドが 1 本もない場合は Trim View の切り替えのみ実行
        if (guideCount === 0) {
            toggleTrimView();
            return;
        }

        previewLayer = createPreviewLayer();
        if (!previewLayer) {
            return;
        }

        duplicateGuidesToPreview();

        if (duplicateFailureCount > 0) {
            if (previewLayer) {
                previewCleanupFailed = !cleanupPartialPreview(previewLayer);
            }

            alert(L(LABELS.err_duplicate_aborted) + "\n" +
                L(LABELS.field_guide_count) + guideCount + "\n" +
                L(LABELS.field_success_count) + duplicateSuccessCount + "\n" +
                L(LABELS.field_failure_count) + duplicateFailureCount +
                (previewCleanupFailed ? "\n" + L(LABELS.err_cleanup_failed) : ""));
            return;
        }

        if (duplicateSuccessCount === 0) {
            alert(L(LABELS.err_preview_failed) + "\n" +
                L(LABELS.field_guide_count) + guideCount + "\n" +
                L(LABELS.field_success_count) + duplicateSuccessCount + "\n" +
                L(LABELS.field_failure_count) + duplicateFailureCount);
            return;
        }

        makeLayerTemplate(previewLayer);
        toggleTrimView();
    } finally {
        try {
            if (docSelection !== null) {
                doc.selection = docSelection;
            }
        } catch (e) { }

        try {
            if (originalActiveLayer) {
                doc.activeLayer = originalActiveLayer;
            }
        } catch (e) { }

        // 元のアイテムのロック状態を復元
        for (i = 0; i < guidesToRelock.length; i++) {
            try {
                guidesToRelock[i].locked = true;
            } catch (e) { }
        }
    }

})();