/*

### 概要

- ドキュメント内の埋め込みラスター画像をPSDに書き出し、リンク画像に置き換えます。
- 実行時のダイアログで、処理対象を「選択している画像のみ」「すべての埋め込み画像」から選べます。
- 書き出し先は、Illustratorドキュメントと同じ階層の「Links」フォルダーです（存在しない場合は自動作成）。
- 書き出すファイル名には、埋め込み前の元のファイル名を使用します。レイヤー名 → 拡張子付きの親グループ名 → XMPマニフェストの順に探し、どれも取得できない場合は image1、image2… の連番になります。
- 同名ファイルがある場合は連番の接尾辞を付けるため、既存ファイルを上書きしません。
- 元の拡大率・回転角・位置を維持したまま、リンク画像として再配置します。

### 注意

- 効果（ドロップシャドウなど）が適用された埋め込み画像の変換は推奨しません。
- 効果はラスタライズされてリンク画像に含まれ、解像度はドキュメントのラスタライズ効果設定ではなく埋め込み画像の解像度に従います。
- 効果によっては、リンク画像の位置が元の埋め込み画像とずれることがあります。
- CMYK／RGB／グレースケール以外のカラースペースの画像はスキップします。

*/

/**
 * @author m1b
 * @discussion https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172
 */

// =========================================
// ユーザー設定
// =========================================
var LINKS_FOLDER_NAME = 'Links';        /* 書き出し先フォルダー名 */
var EXPORT_RESOLUTION = 72;             /* 書き出し解像度（ppi） */
var USE_XMP_NAMES = true;               /* レイヤー名が空のとき、XMPマニフェストの元ファイル名を使う */
var SKIP_RESIZED_IMAGES = false;        /* サイズが一致しない画像は変換せず元に戻す */
var KEEP_EXPORT_DOCUMENT_OPEN = false;  /* 書き出し用の一時ドキュメントを開いたままにする（動作確認用） */

/* 処理対象 */
var SCOPE_SELECTION = 'selection',  /* 選択している画像のみ */
    SCOPE_ALL = 'all';              /* すべての埋め込み画像 */

main();

/**
 * ドキュメントの状態を確認し、処理対象を選んでから実行します。
 * @returns {void}
 */
function main() {

    if (app.documents.length === 0)
        return;

    var sourceDocument = app.activeDocument;

    if (sourceDocument.rasterItems.length === 0) {
        alert('埋め込み画像が見つかりません。');
        return;
    }

    /* 選択範囲に含まれる埋め込み画像（グループの中も探す） */
    var selectedImages = getSelectedRasterImages(sourceDocument);

    var targetScope = showScopeDialog(selectedImages.length);

    if (targetScope == null)
        return;

    var targetImages = (targetScope === SCOPE_SELECTION)
        ? selectedImages
        : collectionToArray(sourceDocument.rasterItems);

    unembedRasterImages(sourceDocument, targetImages);

}

/**
 * 処理対象を選ぶダイアログを表示します。
 * @param {number} selectedImageCount - 選択範囲に含まれる埋め込み画像の件数。
 * @returns {string|null} SCOPE_SELECTION または SCOPE_ALL。キャンセル時は null。
 */
function showScopeDialog(selectedImageCount) {

    var scopeDialog = new Window('dialog', '埋め込み解除');
    scopeDialog.orientation = 'column';
    scopeDialog.alignChildren = 'fill';
    scopeDialog.margins = 16;
    scopeDialog.spacing = 12;

    var targetScopePanel = scopeDialog.add('panel', undefined, '処理対象');
    targetScopePanel.orientation = 'column';
    targetScopePanel.alignChildren = 'left';
    targetScopePanel.margins = [14, 18, 14, 14];
    targetScopePanel.spacing = 8;

    var selectedOnlyRadio = targetScopePanel.add('radiobutton', undefined, '選択している画像のみ（' + selectedImageCount + '件）'),
        allImagesRadio = targetScopePanel.add('radiobutton', undefined, 'すべての埋め込み画像');

    /* 選択範囲に埋め込み画像が無ければ選べない */
    selectedOnlyRadio.enabled = (selectedImageCount > 0);
    selectedOnlyRadio.value = (selectedImageCount > 0);
    allImagesRadio.value = (selectedImageCount === 0);

    var dialogButtonGroup = scopeDialog.add('group');
    dialogButtonGroup.alignment = 'right';
    dialogButtonGroup.add('button', undefined, 'キャンセル', { name: 'cancel' });
    dialogButtonGroup.add('button', undefined, 'OK', { name: 'ok' });

    if (scopeDialog.show() !== 1)
        return null;

    return selectedOnlyRadio.value ? SCOPE_SELECTION : SCOPE_ALL;

}

/**
 * 選択範囲に含まれる埋め込み画像を、グループの中も辿って集めます。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @returns {Array} 選択されている埋め込み画像の配列。
 */
function getSelectedRasterImages(sourceDocument) {

    var foundImages = [];

    collectRasterImages(sourceDocument.selection, foundImages);

    return foundImages;
}

/**
 * ページアイテムから埋め込み画像を再帰的に集めます。
 * @param {Array} pageItems - 探索するページアイテム。
 * @param {Array} foundImages - 見つかった画像を追加する配列。
 * @returns {void}
 */
function collectRasterImages(pageItems, foundImages) {

    for (var i = 0; i < pageItems.length; i++) {

        if (pageItems[i].typename === 'RasterItem')
            foundImages.push(pageItems[i]);

        else if (pageItems[i].typename === 'GroupItem')
            collectRasterImages(pageItems[i].pageItems, foundImages);
    }
}

/**
 * Illustratorのコレクションを配列に変換します。
 * @param {Object} collection - Illustratorのコレクション。
 * @returns {Array} コレクションの要素の配列。
 */
function collectionToArray(collection) {

    var items = [];

    for (var i = 0; i < collection.length; i++)
        items.push(collection[i]);

    return items;
}

/**
 * 指定した埋め込み画像をPSDに書き出し、リンク画像に置き換えます。
 * @author m1b
 * @version 2022-07-23
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @param {Array} targetImages - 変換する埋め込み画像の配列。
 * @returns {void}
 */
function unembedRasterImages(sourceDocument, targetImages) {

    var targetImageCount = targetImages.length,
        convertedImageUUIDs = [],
        reportMessages = [],
        convertedCount = 0;

    var exportFolder = getExportFolder(sourceDocument);

    /* 名前はドキュメント全体で解決する（マニフェストとの突き合わせに全画像が必要） */
    var exportBaseNames = resolveExportBaseNames(sourceDocument, targetImages);

    var previousInteractionLevel = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    /* 変換した画像は最後にまとめて削除する（処理中に削除すると残りの参照が失われる） */
    for (var i = targetImageCount - 1; i >= 0; i--) {

        var embeddedImage = targetImages[i],
            imageColorSpace = embeddedImage.imageColorSpace,
            imageName = exportBaseNames[i];

        if (!isSupportedColorSpace(imageColorSpace)) {
            reportMessages.push('「' + imageName + '」は未対応のカラースペースのためスキップしました。(' + imageColorSpace + ')');
            continue;
        }

        var scaleAndRotation = getScaleAndRotation(embeddedImage),
            exportedFile;

        try {
            exportedFile = exportEmbeddedImage(embeddedImage, exportFolder, imageName, scaleAndRotation);
        } catch (err) {
            reportMessages.push('「' + imageName + '」の書き出しに失敗しました。(' + err.message + ')');
            continue;
        }

        if (!exportedFile.exists) {
            reportMessages.push('「' + imageName + '」の書き出しに失敗しました。');
            continue;
        }

        var linkedImage = placeLinkedImage(embeddedImage, exportedFile, scaleAndRotation);

        /* サイズが変わっている場合は、ドロップシャドウなどの効果が適用されている可能性が高い */
        if (!hasSameSize(embeddedImage, linkedImage)) {

            if (SKIP_RESIZED_IMAGES) {
                reportMessages.push('「' + imageName + '」はサイズを正しく再現できないため変換しませんでした。効果やアピアランスを解除してから再実行してください。');
                exportedFile.remove();
                linkedImage.remove();
                continue;
            }

            reportMessages.push('警告：「' + imageName + '」は効果などの影響でサイズが変化しています。');
        }

        convertedImageUUIDs.push(embeddedImage.uuid);
        convertedCount++;

    }

    removeImagesByUUID(sourceDocument, convertedImageUUIDs);

    app.userInteractionLevel = previousInteractionLevel;
    app.redraw();

    var summaryMessage = targetImageCount + '件中' + convertedCount + '件の埋め込み画像をリンクに変換しました。';

    if (reportMessages.length > 0)
        summaryMessage += '\n' + reportMessages.join('\n');

    alert(summaryMessage);

}

/**
 * 書き出し先フォルダーを返します。無い場合は作成します。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @returns {Folder} 書き出し先フォルダー。
 */
function getExportFolder(sourceDocument) {

    var exportFolder = Folder(sourceDocument.fullName.parent.fsName + '/' + LINKS_FOLDER_NAME + '/');

    if (!exportFolder.exists)
        exportFolder.create();

    return exportFolder;
}

/**
 * 処理対象それぞれの書き出し名を決めます。
 * 名前はレイヤー名 → 拡張子付きの親グループ名 → XMPマニフェスト → 連番 の順に探します。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @param {Array} targetImages - 変換する埋め込み画像の配列。
 * @returns {Array} targetImages と同じ並びの、拡張子を除いた書き出し名。
 */
function resolveExportBaseNames(sourceDocument, targetImages) {

    var allImages = sourceDocument.rasterItems,
        imageNames = [],
        i;

    /* 画像自身から取得できる名前（拡張子付きのまま） */
    for (i = 0; i < allImages.length; i++)
        imageNames.push(getImageNameFromItem(allImages[i]));

    if (USE_XMP_NAMES)
        fillNamesFromManifest(sourceDocument, imageNames);

    /* 選択している画像だけを処理する場合もあるため、uuidで突き合わせる */
    var nameByUUID = {};

    for (i = 0; i < allImages.length; i++)
        nameByUUID[allImages[i].uuid] = toSafeBaseName(imageNames[i], 'image' + (i + 1));

    var exportBaseNames = [];

    for (i = 0; i < targetImages.length; i++)
        exportBaseNames.push(nameByUUID[targetImages[i].uuid] || ('image' + (i + 1)));

    return exportBaseNames;
}

/**
 * 埋め込み画像自身が持つ名前を返します。
 * レイヤー名、無ければ拡張子付きの親グループ名（配置時にファイル名が残ることがある）。
 * @param {RasterItem} embeddedImage - 対象の埋め込み画像。
 * @returns {string} 拡張子付きの名前。取得できない場合は空文字。
 */
function getImageNameFromItem(embeddedImage) {

    if (embeddedImage.name)
        return embeddedImage.name;

    var parentItem = embeddedImage.parent;

    if (parentItem == undefined || parentItem.typename !== 'GroupItem')
        return '';

    var parentName = parentItem.name || '';

    return /\.[a-z][a-z0-9]{1,4}\s*$/i.test(parentName) ? parentName : '';
}

/**
 * 名前が取得できなかった画像の名前を、XMPマニフェストの元ファイル名で補います。
 * 割り当ては、候補の件数が名前未定の画像数と一致するときだけ行い、
 * 一致しない場合は取り違えを避けて何もしません。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @param {Array} imageNames - 画像ごとの名前。空文字の要素が埋められます。
 * @returns {void}
 */
function fillNamesFromManifest(sourceDocument, imageNames) {

    var manifestNames = getManifestFileNames(sourceDocument);

    if (manifestNames.length === 0)
        return;

    /* すでに名前が判明している画像は、その名前を候補から除く */
    var knownNames = {},
        missingCount = 0,
        i;

    for (i = 0; i < imageNames.length; i++) {

        if (imageNames[i] === '') {
            missingCount++;
            continue;
        }

        knownNames[imageNames[i].toLowerCase()] = true;
    }

    if (missingCount === 0)
        return;

    var remainingNames = [];

    for (i = 0; i < manifestNames.length; i++)
        if (!knownNames[manifestNames[i].toLowerCase()])
            remainingNames.push(manifestNames[i]);

    /* 件数が一致しないときは、取り違えを避けるため使わない */
    if (remainingNames.length !== missingCount)
        return;

    var nameIndex = 0;

    for (i = 0; i < imageNames.length; i++)
        if (imageNames[i] === '')
            imageNames[i] = remainingNames[nameIndex++];
}

/**
 * ドキュメントのXMPマニフェストから、埋め込み前の元ファイル名を出現順に取得します。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @returns {Array} 重複を除いた元ファイル名。取得できない場合は空配列。
 */
function getManifestFileNames(sourceDocument) {

    var manifestNames = [],
        foundNames = {},
        filePaths;

    try {
        var documentXMP = new XML(sourceDocument.XMPString);

        /* 埋め込み参照のみを対象にし、取得できなければすべての参照を見る */
        filePaths = documentXMP.xpath('//stMfs:reference/stRef:filePath');

        if (filePaths == null || filePaths.length() === 0)
            filePaths = documentXMP.xpath('//stRef:filePath');

    } catch (err) {
        return manifestNames;
    }

    if (filePaths == null)
        return manifestNames;

    for (var i = 0; i < filePaths.length(); i++) {

        var fileName = decodeURI(String(filePaths[i])).replace(/^.*[\/\\]/, ''),
            nameKey = fileName.toLowerCase();

        if (fileName === '' || foundNames[nameKey])
            continue;

        foundNames[nameKey] = true;
        manifestNames.push(fileName);
    }

    return manifestNames;
}

/**
 * 名前から拡張子と使用できない文字を取り除き、ファイル名として整えます。
 * @param {string} fileName - 元の名前（拡張子付きでも可）。
 * @param {string} fallbackName - 名前が空になる場合に使う代替名。
 * @returns {string} ファイル名に使用できる文字列。
 */
function toSafeBaseName(fileName, fallbackName) {

    /* 拡張子と前後の空白を除去（「2026.07.27」のような名前を壊さないよう英字始まりの2〜5文字に限定） */
    var baseName = (fileName || '').replace(/\.[a-z][a-z0-9]{1,4}$/i, '').replace(/^\s+|\s+$/g, '');

    /* ファイル名に使えない文字を置換 */
    baseName = baseName.replace(/[\\\/:*?"<>|]/g, '_');

    return baseName === '' ? fallbackName : baseName;
}

/**
 * 埋め込み画像を一時ドキュメントへ複製し、等倍・回転なしの状態でPSDに書き出します。
 * @param {RasterItem} embeddedImage - 書き出す埋め込み画像。
 * @param {Folder} exportFolder - 書き出し先フォルダー。
 * @param {string} imageName - 拡張子を除いたファイル名。
 * @param {Object} scaleAndRotation - getScaleAndRotation の戻り値。
 * @returns {File} 書き出したPSDファイル。
 */
function exportEmbeddedImage(embeddedImage, exportFolder, imageName, scaleAndRotation) {

    var imageColorSpace = embeddedImage.imageColorSpace,
        exportDocument = createExportDocument(imageName, imageColorSpace);

    var workingImage = embeddedImage.duplicate(exportDocument.layers[0], ElementPlacement.PLACEATBEGINNING);

    /* 拡大率100%、回転角0°に戻してからアートボードを合わせる */
    var transformMatrix = app.getRotationMatrix(-scaleAndRotation.rotation);
    transformMatrix = app.concatenateScaleMatrix(transformMatrix, 100 / scaleAndRotation.scaleX * 100, 100 / scaleAndRotation.scaleY * 100);
    workingImage.transform(transformMatrix, true, true, true, true, true);
    workingImage.position = [0, workingImage.height];
    exportDocument.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];

    var exportFilePath = getNonOverwritingFilePath(exportFolder.fsName + '/' + imageName + '.psd');

    try {
        return exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, EXPORT_RESOLUTION);
    } finally {
        if (!KEEP_EXPORT_DOCUMENT_OPEN)
            exportDocument.close(SaveOptions.DONOTSAVECHANGES);
    }
}

/**
 * 書き出したファイルを、元の埋め込み画像と同じ拡大率・回転角・位置でリンク配置します。
 * @param {RasterItem} embeddedImage - 元の埋め込み画像。
 * @param {File} linkedFile - リンクするファイル。
 * @param {Object} scaleAndRotation - getScaleAndRotation の戻り値。
 * @returns {PlacedItem} 配置したリンク画像。
 */
function placeLinkedImage(embeddedImage, linkedFile, scaleAndRotation) {

    var originalPosition = embeddedImage.position,
        linkedImage = embeddedImage.layer.placedItems.add();

    linkedImage.file = linkedFile;

    var transformMatrix = app.getScaleMatrix(scaleAndRotation.scaleX, scaleAndRotation.scaleY);
    transformMatrix = app.concatenateRotationMatrix(transformMatrix, scaleAndRotation.rotation);
    linkedImage.transform(transformMatrix, true, true, true, true, true);

    linkedImage.move(embeddedImage, ElementPlacement.PLACEAFTER);
    linkedImage.position = originalPosition;

    return linkedImage;
}

/**
 * 2つのページアイテムの幅と高さが一致するかを返します。
 * @param {PageItem} imageA - 比較するページアイテム。
 * @param {PageItem} imageB - 比較するページアイテム。
 * @returns {boolean} 幅と高さが一致する場合は true。
 */
function hasSameSize(imageA, imageB) {
    return roundTo(imageA.width, 2) === roundTo(imageB.width, 2)
        && roundTo(imageA.height, 2) === roundTo(imageB.height, 2);
}

/**
 * このスクリプトで扱えるカラースペースかどうかを返します。
 * @param {ImageColorSpace} imageColorSpace - 判定するカラースペース。
 * @returns {boolean} CMYK／RGB／グレースケールの場合は true。
 */
function isSupportedColorSpace(imageColorSpace) {
    return imageColorSpace == ImageColorSpace.CMYK
        || imageColorSpace == ImageColorSpace.RGB
        || imageColorSpace == ImageColorSpace.GrayScale;
}

/**
 * UUIDが一致する埋め込み画像をドキュメントから削除します。
 * @param {Document} sourceDocument - 対象のIllustratorドキュメント。
 * @param {Array} imageUUIDs - 削除する画像のuuidの配列。
 * @returns {void}
 */
function removeImagesByUUID(sourceDocument, imageUUIDs) {

    for (var i = 0; i < imageUUIDs.length; i++) {

        var rasterImages = sourceDocument.rasterItems;

        for (var j = 0; j < rasterImages.length; j++) {
            if (rasterImages[j].uuid === imageUUIDs[i]) {
                rasterImages[j].remove();
                break;
            }
        }
    }
}

/**
 * ドキュメントをPSDとして書き出します。
 * @param {Document} exportDocument - 書き出すドキュメント。
 * @param {string} exportFilePath - 書き出し先のパス。
 * @param {ImageColorSpace} imageColorSpace - 画像のカラースペース。
 * @param {number} resolution - 書き出し解像度（ppi）。
 * @returns {File} 書き出したPSDファイル。
 */
function exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, resolution) {

    var exportedFile = File(exportFilePath),
        psdOptions = new ExportOptionsPhotoshop();

    psdOptions.antiAliasing = false;
    psdOptions.artBoardClipping = true;
    psdOptions.imageColorSpace = imageColorSpace;
    psdOptions.editableText = false;
    psdOptions.flatten = true;
    psdOptions.maximumEditability = false;
    psdOptions.resolution = (resolution || 72);
    psdOptions.warnings = false;
    psdOptions.writeLayers = false;

    exportDocument.exportFile(exportedFile, ExportType.PHOTOSHOP, psdOptions);

    return exportedFile;
}

/**
 * PlacedItemまたはRasterItemの拡大率と回転角を返します。
 * @author m1b
 * @version 2022-07-23
 * @param {PlacedItem|RasterItem} placedOrRasterImage - Illustratorの画像アイテム。
 * @returns {Object} { scaleX: 拡大率X%, scaleY: 拡大率Y%, rotation: 回転角° }
 */
function getScaleAndRotation(placedOrRasterImage) {

    /* RasterItemは行列のY方向が反転している */
    var placedItemFlip = (placedOrRasterImage.typename === 'PlacedItem') ? 1 : -1,
        rotationAngle = getAccumulatedRotation(placedOrRasterImage);

    var unrotatedMatrix = app.concatenateRotationMatrix(placedOrRasterImage.matrix, rotationAngle * placedItemFlip);

    return {
        scaleX: unrotatedMatrix.mValueA * 100,
        scaleY: unrotatedMatrix.mValueD * -100 * placedItemFlip,
        rotation: rotationAngle
    };
}

/**
 * 画像アイテムに蓄積された回転角を度で返します。
 * @param {PlacedItem|RasterItem} placedOrRasterImage - Illustratorの画像アイテム。
 * @returns {number} 回転角（度）。取得できない場合は 0。
 */
function getAccumulatedRotation(placedOrRasterImage) {

    var imageTags = placedOrRasterImage.tags;

    for (var i = 0; i < imageTags.length; i++) {
        if (imageTags[i].name === 'BBAccumRotation')
            return imageTags[i].value * 180 / Math.PI;
    }

    return 0;
}

/**
 * `value` を小数点以下 `places` 桁に丸めます。
 * @param {number} value - 丸める数値。
 * @param {number} places - 小数点以下の桁数。負の値も指定できます。
 * @returns {number} 丸めた数値。
 */
function roundTo(value, places) {
    var factor = Math.pow(10, places || 3);
    return Math.round(value * factor) / factor;
}

/**
 * オプションを指定して書き出し用の新規ドキュメントを作成します。
 * @author m1b
 * @version 2022-07-21
 * @param {string} documentTitle - ドキュメントの名前。
 * @param {ImageColorSpace} imageColorSpace - 画像のカラースペース。
 * @returns {Document} 作成したドキュメント。
 */
function createExportDocument(documentTitle, imageColorSpace) {

    var documentPreset = new DocumentPreset(),
        documentPresetType;

    documentPreset.title = documentTitle;
    documentPreset.width = 1000;
    documentPreset.height = 1000;

    if (imageColorSpace == ImageColorSpace.RGB) {
        documentPresetType = DocumentPresetType.BasicRGB;
        documentPreset.colorMode = DocumentColorSpace.RGB;
    }

    else {
        documentPresetType = DocumentPresetType.BasicCMYK;
        documentPreset.colorMode = DocumentColorSpace.CMYK;
    }

    return app.documents.addDocument(documentPresetType, documentPreset);

}

/**
 * 既存ファイルを上書きしないパスを返します。
 * 上書きを避けるため、連番の接尾辞を付けます。
 * @author m1b
 * @version 2022-07-20
 *
 * 例えば `myFile.txt` がすでに存在する場合は
 * `myFile(1).txt` を返し、それも存在する場合は
 * `myFile(2).txt` を返します。
 *
 * @param {string} filePath - 調べるパス。
 * @returns {string} 既存ファイルを上書きしないパス。
 */
function getNonOverwritingFilePath(filePath) {

    var suffixIndex = 1,
        pathParts = filePath.split(/(\.[^\.]+)$/);

    while (File(filePath).exists)
        filePath = pathParts[0] + '(' + (++suffixIndex) + ')' + pathParts[1];

    return filePath;
}
