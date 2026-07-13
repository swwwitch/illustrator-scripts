#target illustrator

/*

# スクリプトの概要

アクティブなアートボードを PNG24 形式で書き出します。

- 解像度 200%、アートボードでクリップ、アンチエイリアスあり
- 背景は白に固定（透過なし）
- 保存先はドキュメントと同じフォルダ、ファイル名は「ドキュメント名.png」形式
- 書き出し中は「Guides Preview for Trim View」レイヤーを一時的に非表示にし、完了後に再表示
- 書き出し後、Mac では保存フォルダを Finder で自動的に開く
- 最後に書き出したファイルのパス（exportFolder + exportFileName）を return

作成日：2025-04-22

*/

main();

/* メイン処理：アクティブアートボードを PNG 書き出し / Main: export the active artboard as PNG */
function main() {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var exportFolder = doc.fullName.parent;
    var documentBaseName = doc.name.replace(/\.ai$/i, "");

    var hiddenGuidesLayer = hideLayerByName(doc, "Guides Preview for Trim View");

    var pngExportOptions = new ExportOptionsPNG24();
    pngExportOptions.artBoardClipping = true;
    pngExportOptions.antiAliasing = true;
    pngExportOptions.transparency = false; // 背景：白
    pngExportOptions.horizontalScale = 200;
    pngExportOptions.verticalScale = 200;

    var exportFileName = documentBaseName + ".png";
    var exportFile = new File(exportFolder + "/" + exportFileName);

    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    try {
        doc.exportFile(exportFile, ExportType.PNG24, pngExportOptions);

        if (Folder.fs === "Macintosh") {
            exportFolder.execute(); // Finder で保存先を開く
        }
    } catch (e) {
        alert("書き出し中にエラーが発生しました：\n" + e.message);
    } finally {
        if (hiddenGuidesLayer) {
            hiddenGuidesLayer.visible = true; // 書き出し後に再表示
        }
    }

    app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

    return exportFolder + "/" + exportFileName;
}

/* 指定名のレイヤーを非表示にして返す（なければ null） / Hide the layer with the given name and return it (null if absent) */
function hideLayerByName(doc, layerName) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            doc.layers[i].visible = false;
            return doc.layers[i];
        }
    }
    return null;
}