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
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

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
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            saveFirst: { ja: "先にドキュメントを保存してください。", en: "Save the document first." },
            exportError: {
                ja: "書き出し中にエラーが発生しました：\n",
                en: "An error occurred during export:\n"
            }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "alert.saveFirst" のようなドット区切りのキー
     * @returns {string} 現在のUI言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

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
        /* 一度も保存していないドキュメントは保存先が定まらないため中止（fullName が例外になることがある）/ Abort when the document has never been saved (fullName may throw) */
        var exportFolder = null;
        try {
            exportFolder = doc.fullName.parent;
        } catch (e) {}
        if (!exportFolder || !exportFolder.exists) {
            alert(getLabel("alert.saveFirst"));
            return;
        }
        var documentBaseName = doc.name.replace(/\.ai$/i, "");
        var exportFile = new File(exportFolder.fsName + "/" + documentBaseName + ".png");

        var hiddenGuidesLayer = hideVisibleLayerByName(doc, GUIDES_LAYER_NAME);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        /* 書き出しの失敗は知らせ、隠したレイヤーと警告表示の設定は必ず戻す / Report export errors and always restore the layer and alert level */
        try {
            doc.exportFile(exportFile, ExportType.PNG24, createPngExportOptions());
            if (Folder.fs === "Macintosh") {
                exportFolder.execute(); /* Finder で保存先を開く / Open the destination in Finder */
            }
        } catch (e) {
            alert(getLabel("alert.exportError") + e.message);
        } finally {
            if (hiddenGuidesLayer) {
                hiddenGuidesLayer.visible = true; /* 書き出し後に再表示 / Show it again after export */
            }
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        }
    }

    /**
     * 背景白・指定倍率の PNG24 書き出し設定を作る
     * @returns {ExportOptionsPNG24} 書き出し設定
     */
    function createPngExportOptions() {
        var pngExportOptions = new ExportOptionsPNG24();
        pngExportOptions.artBoardClipping = true;
        pngExportOptions.antiAliasing = true;
        pngExportOptions.transparency = false; /* 背景：白 / White background */
        pngExportOptions.horizontalScale = EXPORT_SCALE;
        pngExportOptions.verticalScale = EXPORT_SCALE;
        return pngExportOptions;
    }

    /**
     * 指定名のレイヤーが表示中なら非表示にして返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 非表示にしたレイヤー（見つからない、または元から非表示なら null）
     */
    function hideVisibleLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (candidateLayer.name === layerName) {
                /* 元から非表示なら書き出し後も非表示のまま / Leave an already hidden layer hidden */
                if (!candidateLayer.visible) return null;
                candidateLayer.visible = false;
                return candidateLayer;
            }
        }
        return null;
    }

    main();

})();
