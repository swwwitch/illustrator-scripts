#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなアートボードを、解像度200%・背景白のPNG24として、ドキュメントと同じフォルダーへ書き出します。
書き出し中は「Guides Preview for Trim View」レイヤーを一時的に非表示にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export200.md

### Overview

Exports the active artboard as a PNG24 at 200% scale on a white background, into the same folder as the document.
The "Guides Preview for Trim View" layer is hidden while the export runs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export200.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "export200";                    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export200.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export200.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var EXPORT_SCALE      = 200;                            /* 書き出し倍率（%）/ Export scale (%) */
    var GUIDES_LAYER_NAME = "Guides Preview for Trim View"; /* 書き出し中に隠すレイヤー / Layer hidden during export */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * アクティブなアートボードを背景白の PNG24 としてドキュメントと同じフォルダーへ書き出す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;
        var exportFolder = doc.fullName.parent;
        var documentBaseName = doc.name.replace(/\.ai$/i, "");

        var hiddenGuidesLayer = hideLayerByName(doc, GUIDES_LAYER_NAME);

        var pngExportOptions = new ExportOptionsPNG24();
        pngExportOptions.artBoardClipping = true;
        pngExportOptions.antiAliasing = true;
        pngExportOptions.transparency = false; /* 背景：白 / White background */
        pngExportOptions.horizontalScale = EXPORT_SCALE;
        pngExportOptions.verticalScale = EXPORT_SCALE;

        var exportFile = new File(exportFolder + "/" + documentBaseName + ".png");

        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        /* 書き出しの失敗は知らせ、隠したレイヤーは必ず戻す / Report export errors and always restore the hidden layer */
        try {
            doc.exportFile(exportFile, ExportType.PNG24, pngExportOptions);

            if (Folder.fs === "Macintosh") {
                exportFolder.execute(); /* Finder で保存先を開く / Open the destination in Finder */
            }
        } catch (e) {
            alert("書き出し中にエラーが発生しました：\n" + e.message);
        } finally {
            if (hiddenGuidesLayer) {
                hiddenGuidesLayer.visible = true; /* 書き出し後に再表示 / Show it again after export */
            }
        }

        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
    }

    /**
     * 指定名のレイヤーを非表示にして返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 非表示にしたレイヤー（見つからなければ null）
     */
    function hideLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (candidateLayer.name === layerName) {
                candidateLayer.visible = false;
                return candidateLayer;
            }
        }
        return null;
    }

    main();

})();
