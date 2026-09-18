#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを高解像度でラスタライズし、Figmaの「Copy as PNG」のようにビットマップとしてクリップボードへコピーします。
600dpiでラスタライズしたあと、72ppi相当の偶数整数倍率に調整します。

詳細は README を参照してください。

### Overview

Rasterizes the selection at high resolution and copies it to the clipboard as a bitmap, the way Figma's "Copy as PNG" does.
It rasterizes at 600 dpi, then scales the result to an even integer multiple of 72 ppi.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyAsPngLikeFigma";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-06-03";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigma.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigma.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf5f269788086"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // -------------------------------
    // 日英ラベル定義 Define label
    // -------------------------------
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = getCurrentLang();
    var LABELS = {
        progressTitle: { ja: "処理中...", en: "Processing..." },
        progressStart: { ja: "処理を開始しています...", en: "Starting process..." },
        progressLayer: { ja: "一時レイヤー作成中...", en: "Creating temp layer..." },
        progressDuplicate: { ja: "複製中...", en: "Duplicating..." },
        progressRect: { ja: "ラスタライズ範囲作成中...", en: "Creating raster area..." },
        progressRaster: { ja: "ラスタライズ中...", en: "Rasterizing..." },
        progressResize: { ja: "拡大中...", en: "Resizing..." },
        progressCopy: { ja: "コピー中...", en: "Copying..." },
        progressClean: { ja: "一時オブジェクト削除中...", en: "Cleaning up..." },
        progressDone: { ja: "完了しました。", en: "Completed." },
        alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        error: { ja: "エラーが発生しました: ", en: "An error occurred: " }
    };

    function main() {
        if (app.documents.length === 0 || app.selection.length === 0) {
            alert(LABELS.alertSelect[uiLang]);
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = app.selection;

        // プログレスバー付きダイアログの作成
        var progressWin = new Window("palette", LABELS.progressTitle[uiLang]);
        progressWin.pbar = progressWin.add("progressbar", [20, 20, 300, 10], 0, 100);
        progressWin.st = progressWin.add("statictext", undefined, LABELS.progressStart[uiLang]);
        progressWin.show();
        function updateProgress(value, text) {
            progressWin.pbar.value = value;
            progressWin.st.text = text;
            progressWin.update();
        }

        try {
            updateProgress(10, LABELS.progressLayer[uiLang]);
            // 一時レイヤーの作成
            var tempLayer = doc.layers.add();
            tempLayer.name = "__TEMP_LAYER__";
            tempLayer.locked = false;
            tempLayer.visible = true;

            updateProgress(20, LABELS.progressDuplicate[uiLang]);
            // オブジェクトを複製して一時グループに追加
            var duplicatedItems = [];
            var tempGroup = tempLayer.groupItems.add();
            for (var i = 0; i < originalSelection.length; i++) {
                var dup = originalSelection[i].duplicate(tempGroup, ElementPlacement.PLACEATEND);
                duplicatedItems.push(dup);
            }

            updateProgress(35, LABELS.progressRect[uiLang]);
            var bounds = tempGroup.visibleBounds;
            var rect = doc.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
            rect.stroked = false;
            rect.filled = false;
            rect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);

            updateProgress(50, LABELS.progressRaster[uiLang]);
            var resolution = 600;
            var options = new RasterizeOptions();
            options.resolution = resolution;
            options.transparency = false;
            options.backgroundBlack = false;
            options.antiAliasing = true;

            var rasterized = doc.rasterize(tempGroup, rect.geometricBounds, options);

            updateProgress(70, LABELS.progressResize[uiLang]);
            // 拡大倍率を「72ppi相当」を基準に偶数の整数に調整
            var baseRatio = (resolution / 72) * 100;
            var resizeRatio = Math.ceil(baseRatio / 2) * 2; // 偶数の整数倍に切り上げ
            rasterized.resize(resizeRatio, resizeRatio);

            updateProgress(85, LABELS.progressCopy[uiLang]);
            app.selection = [rasterized];
            app.executeMenuCommand("copy");

            updateProgress(95, LABELS.progressClean[uiLang]);
            try { rasterized.remove(); } catch (e) {}
            try { rect.remove(); } catch (e) {}
            for (var j = 0; j < duplicatedItems.length; j++) {
                try { duplicatedItems[j].remove(); } catch (e) {}
            }
            try { tempGroup.remove(); } catch (e) {}
            try { tempLayer.remove(); } catch (e) {}

            app.selection = originalSelection;

            updateProgress(100, LABELS.progressDone[uiLang]);
        } catch (err) {
            progressWin.close();
            alert(LABELS.error[uiLang] + err.message);
            return;
        }

        progressWin.close();
    }

    main();

})();
