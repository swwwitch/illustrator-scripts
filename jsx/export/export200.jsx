#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなアートボードを、解像度200%・背景白のPNG24として、ドキュメントと同じフォルダーへ書き出します。
書き出し中は「Guides Preview for Trim View」レイヤーを一時的に非表示にします。

詳細は README を参照してください。

### Overview

Exports the active artboard as a PNG24 at 200% scale on a white background, into the same folder as the document.
The "Guides Preview for Trim View" layer is hidden while the export runs.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "export200";                    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-04-22";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export200.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export200.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

})();
