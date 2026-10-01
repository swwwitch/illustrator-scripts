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
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

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

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    // ファイルビューアで表示（再利用パーツ） / Show in file viewer (reusable)

    /**
     * Path Finder（起動中のとき）か Finder で、フォルダーを開くかファイルを選択して表示する。
     * 補助アプリ /Applications/OpenInFileViewer.app に一時ファイルでパスを渡して起動する。
     * 補助アプリは illustrator-scripts の helpers/OpenInFileViewer.applescript から作る
     * @param {File|Folder} targetItem - 開くフォルダーか、選択して表示するファイル
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function openInFileViewer(targetItem) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var viewerAppPath = "/Applications/OpenInFileViewer.app";
        var pathFilePath = "/tmp/open_in_file_viewer_path.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(viewerAppPath).exists && !new File(viewerAppPath).exists) return false;

        var pathFile = new File(pathFilePath);
        var written = false;
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (pathFile.open("w")) {
                /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
                written = pathFile.write(targetItem.fsName);
            }
        } catch (e) {
        } finally {
            try { pathFile.close(); } catch (closeError) {}
        }
        return written && new File(viewerAppPath).execute();
    }

    // ファイルビューアで表示（再利用パーツ）ここまで / End of the reusable file viewer

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
                /* Path Finder（起動中のとき）か Finder で保存先を開く / Open the destination in Path Finder or Finder */
                if (!openInFileViewer(exportFolder)) exportFolder.execute();
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
