#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


// ===== Script info =====
var SCRIPT_NAME = 'ImageLinkManager';
var SCRIPT_VERSION = '1.0.0';


// 【概要 / Summary】
// 配置画像（PlacedItem）を埋め込み（embed）に変換する／解除（埋め込み解除：RasterItem をリンク配置に戻す）するスクリプトです。
// ダイアログ上部のモードで「埋め込み / 解除」を切り替え、対応パネルのみ有効化（それ以外はディム表示）します。
// それぞれのパネルで処理対象を選択できます：
// ・選択した画像のみ（選択オブジェクト配下を再帰探索）
// ・すべての画像（ドキュメント内）
//
// 【埋め込み】
// PlacedItem を embed() します。PSD（.psd）の場合は embed() 後に ungroup を試行します。
// ただし親グループがクリッピンググループ（clipped === true）の場合は ungroup を行いません。
//
// 【解除】
// 埋め込み画像（RasterItem）を PSD に書き出し、リンク配置（PlacedItem）へ置き換えます。
// 可能な範囲で位置・回転・拡大率を合わせます（効果が適用されている RasterItem では差異が出る場合があります）。
//
// 実行後、選択を解除します。
// 更新日 / Updated: 2025-12-20

function showDialog() {
    var dlg = new Window('dialog', SCRIPT_NAME + ' v' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 10];

    // ===== モード選択（上部）/ Mode (top) =====
    var modeWrap = dlg.add('group');
    modeWrap.orientation = 'row';
    modeWrap.alignment = 'center';

    var modeGroup = modeWrap.add('group');
    modeGroup.orientation = 'row';
    modeGroup.alignChildren = ['left', 'center'];

    var rbModeEmbed = modeGroup.add('radiobutton', undefined, '埋め込み');
    var rbModeRelease = modeGroup.add('radiobutton', undefined, '解除');
    rbModeEmbed.value = true; // デフォルト

    // ===== 埋め込み / Embed =====
    var panelEmbed = dlg.add('panel', undefined, '埋め込み');
    panelEmbed.orientation = 'column';
    panelEmbed.alignChildren = 'left';
    panelEmbed.margins = [15, 20, 15, 10];

    var rbEmbedSelection = panelEmbed.add('radiobutton', undefined, '選択した配置画像のみ');
    var rbEmbedAll = panelEmbed.add('radiobutton', undefined, 'すべての配置画像');
    rbEmbedSelection.value = true; // デフォルト

    // 選択がない場合は「すべての配置画像」を自動選択
    var hasSelection = false;
    try {
        hasSelection = (app.documents.length > 0 && app.activeDocument.selection && app.activeDocument.selection.length > 0);
    } catch (e) {
        hasSelection = false;
    }
    if (!hasSelection) {
        rbEmbedSelection.value = false;
        rbEmbedAll.value = true;
    }

    // ===== 解除 / Release =====
    var panelRelease = dlg.add('panel', undefined, '解除');
    panelRelease.orientation = 'column';
    panelRelease.alignChildren = 'left';
    panelRelease.margins = [15, 20, 15, 10];

    var rbReleaseSelection = panelRelease.add('radiobutton', undefined, '選択した画像のみ');
    var rbReleaseAll = panelRelease.add('radiobutton', undefined, 'すべての画像');
    rbReleaseSelection.value = true; // デフォルト

    // 選択がない場合は「すべての画像」を自動選択（解除）
    if (!hasSelection) {
        rbReleaseSelection.value = false;
        rbReleaseAll.value = true;
    }

    // UI：パネルのディム切替
    function updatePanelState() {
        panelEmbed.enabled = rbModeEmbed.value;
        panelRelease.enabled = rbModeRelease.value;
    }
    rbModeEmbed.onClick = updatePanelState;
    rbModeRelease.onClick = updatePanelState;
    updatePanelState();

    // ===== Buttons (OK right) =====
    var btnGroup = dlg.add('group');
    btnGroup.alignment = 'right';
    var cancelBtn = btnGroup.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var okBtn = btnGroup.add('button', undefined, 'OK', { name: 'ok' });

    okBtn.onClick = function () {
        dlg.close(1);
    };
    cancelBtn.onClick = function () {
        dlg.close(0);
    };

    var result = dlg.show();
    if (result !== 1) return null;

    return {
        mode: rbModeEmbed.value ? 'embed' : 'release',
        embedMode: rbEmbedAll.value ? 'all' : 'selection',
        releaseMode: rbReleaseAll.value ? 'all' : 'selection'
    };
}

// ===== Collectors =====
function collectPlacedItemsFromSelection(selection) {
    var results = [];

    function walk(item) {
        if (!item) return;
        if (item.typename === 'PlacedItem') {
            results.push(item);
            return;
        }
        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        }
    }

    for (var i = 0; i < selection.length; i++) {
        walk(selection[i]);
    }

    return results;
}

function collectAllPlacedItems(doc) {
    var results = [];
    var items = doc.placedItems;
    for (var i = 0; i < items.length; i++) {
        results.push(items[i]);
    }
    return results;
}

function collectRasterItemsFromSelection(selection) {
    var results = [];

    function isEmbeddedPlacedItem(it) {
        if (!it || it.typename !== 'PlacedItem') return false;
        try {
            return (it.file == null);
        } catch (e) {
            return true;
        }
    }

    function walk(item) {
        if (!item) return;

        if (item.typename === 'RasterItem') {
            results.push(item);
            return;
        }

        if (isEmbeddedPlacedItem(item)) {
            results.push(item);
            return;
        }

        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        }
    }

    for (var i = 0; i < selection.length; i++) {
        walk(selection[i]);
    }

    return results;
}

function collectAllRasterItems(doc) {
    var results = [];

    // RasterItem（埋め込み画像）
    try {
        var r = doc.rasterItems;
        for (var i = 0; i < r.length; i++) {
            results.push(r[i]);
        }
    } catch (e) { }

    // 念のため：埋め込み後も PlacedItem のまま残るケース
    try {
        var p = doc.placedItems;
        for (var j = 0; j < p.length; j++) {
            try {
                if (p[j].file == null) {
                    results.push(p[j]);
                }
            } catch (e2) {
                // file 参照で例外の場合も embedded 扱い
                results.push(p[j]);
            }
        }
    } catch (e3) { }

    return results;
}

function selectionHasLinkedPlacedItem(selection) {
    function isLinkedPlacedItem(it) {
        if (!it || it.typename !== 'PlacedItem') return false;
        try {
            return (it.file != null && it.file.exists);
        } catch (e) {
            // file 参照で例外の場合は「リンク」とは断定しない
            return false;
        }
    }

    function walk(item) {
        if (!item) return false;
        if (isLinkedPlacedItem(item)) return true;
        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                if (walk(item.pageItems[i])) return true;
            }
        }
        return false;
    }

    for (var i = 0; i < selection.length; i++) {
        if (walk(selection[i])) return true;
    }
    return false;
}

// ===== Embed =====
function embedPlacedItems(placedItems) {
    var count = 0;

    for (var i = 0; i < placedItems.length; i++) {
        var item = placedItems[i];
        if (!item) continue;

        // 親がクリップグループかどうか
        var parent = item.parent;
        var isClipGroup = false;
        if (parent && parent.typename === 'GroupItem') {
            isClipGroup = parent.clipped;
        }

        // PSDファイルかどうか
        if (item.file && item.file.name.match(/\.psd$/i)) {
            try {
                item.embed();
                count++;
            } catch (e) {
                continue;
            }

            if (!isClipGroup) {
                try {
                    app.executeMenuCommand('ungroup');
                } catch (e2) { }
            }
        } else {
            try {
                item.embed();
                count++;
            } catch (e3) {
                continue;
            }
        }
    }

    return count;
}

function sanitizeFileBaseName(name) {
    if (!name) return '';
    // strip any path
    name = String(name).replace(/^.*[\\\/]/, '');
    // strip extension
    name = name.replace(/\.[^.]+$/, '');
    // replace invalid characters for mac/windows
    name = name.replace(/[\/:*?"<>|\r\n]+/g, '_');
    name = name.replace(/^\s+|\s+$/g, '');
    if (name.length > 120) name = name.substring(0, 120);
    return name;
}

function getFileNameFromXMPForEmbeddedItem(item) {
    try {
        var xmpStr = item.XMPString;
        if (!xmpStr) return '';

        // dc:title
        var titleMatch = xmpStr.match(/<dc:title>\s*<rdf:Alt>\s*<rdf:li[^>]*>(.*?)<\/rdf:li>/);
        if (titleMatch && titleMatch[1]) return titleMatch[1];

        // tiff:ImageDescription
        var descMatch = xmpStr.match(/tiff:ImageDescription>(.*?)<\/tiff:ImageDescription>/);
        if (descMatch && descMatch[1]) return descMatch[1];

        // photoshop:Source (sometimes contains a path)
        var srcMatch = xmpStr.match(/photoshop:Source="(.*?)"/);
        if (srcMatch && srcMatch[1]) return srcMatch[1];

        return '';
    } catch (e) {
        return '';
    }
}

function getPreferredBaseNameForUnembed(item, fallbackBaseName) {
    var name = '';

    // 1) item.name (レイヤーパネルの名称にファイル名が残っていることが多い)
    try {
        if (item.name && item.name !== '') name = item.name;
    } catch (e1) { }

    // 2) item.file.name（埋め込み後は null が多いが念のため）
    if (!name) {
        try {
            if (item.file && item.file.name) name = item.file.name;
        } catch (e2) { }
    }

    // 3) XMP
    if (!name) {
        name = getFileNameFromXMPForEmbeddedItem(item);
    }

    name = sanitizeFileBaseName(name);
    if (!name) name = fallbackBaseName;

    return name;
}

// ===== Release (Unembed) =====
function releaseRasterItemsToLinks(doc, rasterItems) {
    var previousInteractionLevel = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    var exportErrors = [];
    var counter = 0;
    var linkedCount = 0;

    // export folder
    var exportFolder;
    try {
        exportFolder = Folder(doc.fullName.parent.fsName + '/Links/');
    } catch (e) {
        exportFolder = Folder(Folder.desktop.fsName + '/Links/');
    }
    if (!exportFolder.exists) {
        try { exportFolder.create(); } catch (e2) { }
    }

    try {
        // 後ろから処理（順序変化に強く）
        for (var i = rasterItems.length - 1; i >= 0; i--) {
            var oldImage = rasterItems[i];
            if (!oldImage) continue;

            var fallbackTitle = 'image' + (i + 1);
            var imageTitle = getPreferredBaseNameForUnembed(oldImage, fallbackTitle);

            // colorSpace
            var colorSpace;
            try {
                // RasterItem
                colorSpace = oldImage.imageColorSpace;
            } catch (eCS) {
                // embedded PlacedItem など
                try {
                    colorSpace = (doc.documentColorSpace === DocumentColorSpace.CMYK) ? ImageColorSpace.CMYK : ImageColorSpace.RGB;
                } catch (eCS2) {
                    colorSpace = ImageColorSpace.RGB;
                }
            }

            // script can't handle other colorSpaces
            if (
                colorSpace != ImageColorSpace.CMYK &&
                colorSpace != ImageColorSpace.RGB &&
                colorSpace != ImageColorSpace.GrayScale
            ) {
                exportErrors.push(imageTitle + ' has unsupported color space. (' + colorSpace + ')');
                continue;
            }

            var position;
            try { position = oldImage.position; } catch (ePos) { position = [0, 0]; }

            // get current scale and rotation
            var sr;
            try {
                sr = getLinkScaleAndRotation(oldImage);
            } catch (eSR) {
                sr = [100, 100, 0];
            }
            var scale = [sr[0], sr[1]];
            var rotation = sr[2];

            // move to new document for exporting
            var temp;
            try {
                temp = newDocument(imageTitle, colorSpace, 1000, 1000);
            } catch (eDoc) {
                exportErrors.push('"' + imageTitle + '" failed to create temp doc. (' + eDoc.message + ')');
                continue;
            }

            // duplicate to new document
            var workingImage;
            try {
                workingImage = oldImage.duplicate(temp.layers[0], ElementPlacement.PLACEATBEGINNING);
            } catch (eDup) {
                exportErrors.push('"' + imageTitle + '" failed to duplicate. (' + eDup.message + ')');
                try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose1) { }
                continue;
            }

            // set image to 100% scale 0° rotation and position
            try {
                var tm = app.getRotationMatrix(-rotation);
                tm = app.concatenateScaleMatrix(tm, 100 / scale[0] * 100, 100 / scale[1] * 100);
                workingImage.transform(tm, true, true, true, true, true);
                workingImage.position = [0, workingImage.height];
                temp.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];
            } catch (eTx) { }

            // export
            var file;
            try {
                var path = getFilePathWithOverwriteProtectionSuffix(exportFolder.fsName + '/' + imageTitle + '.psd');
                file = exportAsPSD(temp, path, colorSpace, 72);
            } catch (error) {
                exportErrors.push('"' + imageTitle + '" failed to export. (' + error.message + ')');
            }

            // close temp doc
            try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose2) { }

            if (!file || !file.exists) {
                continue;
            }

            // place link in active layer then move to oldImage position in stack
            var newImage;
            try {
                var targetLayer = doc.activeLayer || oldImage.layer;
                newImage = targetLayer.placedItems.add();

                // まず file を設定（基本）
                newImage.file = file;

                // 環境によっては file 代入だけでは反映されないことがあるため、relink/update を試行
                try {
                    if (newImage.relink) newImage.relink(file);
                } catch (eRelink) { }
                try {
                    if (newImage.update) newImage.update();
                } catch (eUpdate) { }

                // リンクできているか簡易検証
                try {
                    if (newImage.file != null && newImage.file.exists) {
                        linkedCount++;
                    }
                } catch (eCheck) { }

            } catch (eLink) {
                exportErrors.push('"' + imageTitle + '" failed to place link. (' + eLink.message + ')');
                try { if (file && file.exists) file.remove(); } catch (eRm) { }
                continue;
            }

            // scale, rotate, position to match original
            try {
                var tm2 = app.getScaleMatrix(scale[0], scale[1]);
                tm2 = app.concatenateRotationMatrix(tm2, rotation);
                newImage.transform(tm2, true, true, true, true, true);
            } catch (eFit1) { }

            // stacking + position
            try {
                newImage.move(oldImage, ElementPlacement.PLACEAFTER);
            } catch (eMove) { }
            try {
                newImage.position = position;
            } catch (eFit2) { }

            // remove old embedded image
            try {
                oldImage.remove();
            } catch (eDel) {
                exportErrors.push('Warning: Could not remove embedded image for "' + imageTitle + '".');
            }

            counter++;
        }
    } finally {
        app.userInteractionLevel = previousInteractionLevel;
        try { app.redraw(); } catch (eR) { }
    }

    return {
        count: counter,
        linkedCount: linkedCount,
        errors: exportErrors
    };
}

// ===== Helpers (from reference, simplified) =====
function exportAsPSD(doc, path, colorSpace, resolution) {
    var file = File(path);
    var options = new ExportOptionsPhotoshop();

    options.antiAliasing = false;
    options.artBoardClipping = true;
    options.imageColorSpace = colorSpace;
    options.editableText = false;
    options.flatten = true;
    options.maximumEditability = false;
    options.resolution = (resolution || 72);
    options.warnings = false;
    options.writeLayers = false;

    doc.exportFile(file, ExportType.PHOTOSHOP, options);
    return file;
}

function getLinkScaleAndRotation(item) {
    if (item == undefined) return;

    var m = item.matrix;
    var rotatedAmount;
    var unrotatedMatrix;
    var scaledAmount;

    var flipPlacedItem = (item.typename == 'PlacedItem') ? 1 : -1;

    try {
        rotatedAmount = item.tags.getByName('BBAccumRotation').value * 180 / Math.PI;
    } catch (error) {
        rotatedAmount = 0;
    }

    unrotatedMatrix = app.concatenateRotationMatrix(m, rotatedAmount * flipPlacedItem);
    scaledAmount = [unrotatedMatrix.mValueA * 100, unrotatedMatrix.mValueD * -100 * flipPlacedItem];

    return [scaledAmount[0], scaledAmount[1], rotatedAmount];
}

function newDocument(name, colorSpace, width, height) {
    var myDocPreset = new DocumentPreset();
    var myDocPresetType;

    myDocPreset.title = name;
    myDocPreset.width = width || 1000;
    myDocPreset.height = height || 1000;

    if (
        colorSpace == ImageColorSpace.CMYK ||
        colorSpace == ImageColorSpace.GrayScale ||
        colorSpace == ImageColorSpace.DeviceN
    ) {
        myDocPresetType = DocumentPresetType.BasicCMYK;
        myDocPreset.colorMode = DocumentColorSpace.CMYK;
    } else {
        myDocPresetType = DocumentPresetType.BasicRGB;
        myDocPreset.colorMode = DocumentColorSpace.RGB;
    }

    return app.documents.addDocument(myDocPresetType, myDocPreset);
}

function getFilePathWithOverwriteProtectionSuffix(path) {
    var index = 1;
    var parts = path.split(/(\.[^\.]+)$/);

    while (File(path).exists) {
        path = parts[0] + '(' + (++index) + ')' + parts[1];
    }

    return path;
}

// ===== Main =====
function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;

    var dialogResult = showDialog();
    if (!dialogResult) return;

    // モード分岐
    if (dialogResult.mode === 'embed') {
        var placedItems = [];

        if (dialogResult.embedMode === 'selection') {
            if (!doc.selection || doc.selection.length === 0) {
                doc.selection = null;
                return;
            }
            placedItems = collectPlacedItemsFromSelection(doc.selection);
        } else {
            placedItems = collectAllPlacedItems(doc);
        }

        if (placedItems.length === 0) {
            doc.selection = null;
            return;
        }

        var embeddedCount = embedPlacedItems(placedItems);
        if (embeddedCount > 0) {
            alert(embeddedCount + '個の画像を埋め込みました');
        }

        doc.selection = null;
        return;
    }

    // 解除（埋め込み解除）
    var rasterItems = [];

    if (dialogResult.releaseMode === 'selection') {
        if (!doc.selection || doc.selection.length === 0) {
            doc.selection = null;
            return;
        }
        // 選択に「リンク」画像（PlacedItem）が含まれている場合は解除できないため通知
        if (selectionHasLinkedPlacedItem(doc.selection)) {
            alert('リンク画像は解除できません。\n埋め込み画像を選択して実行してください。');
            doc.selection = null;
            return;
        }
        rasterItems = collectRasterItemsFromSelection(doc.selection);
    } else {
        rasterItems = collectAllRasterItems(doc);
    }

    if (rasterItems.length === 0) {
        doc.selection = null;
        return;
    }

    var result = releaseRasterItemsToLinks(doc, rasterItems);
    if (result) {
        if (result.count > 0) {
            var msg = result.count + '個の画像を解除しました';
            if (typeof result.linkedCount === 'number') {
                msg += '\n（リンク作成: ' + result.linkedCount + '）';
                if (result.linkedCount === 0) {
                    msg += '\n\n※リンク作成が確認できませんでした。書き出し先フォルダや権限、ドキュメントの保存状態をご確認ください。';
                }
            }
            if (result.errors && result.errors.length > 0) {
                msg += '\n\n' + result.errors.join('\n');
            }
            alert(msg);
        } else if (result.errors && result.errors.length > 0) {
            alert('解除できませんでした。\n\n' + result.errors.join('\n'));
        }
    }

    doc.selection = null;
}

main();