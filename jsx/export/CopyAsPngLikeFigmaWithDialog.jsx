#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを高解像度でラスタライズし、PNG相当のビットマップとしてクリップボードへコピーします。
解像度・背景色・アンチエイリアス・余白は、実行時のダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigmaWithDialog.md

### Overview

Rasterizes the selection at high resolution and copies it to the clipboard as a PNG-equivalent bitmap.
Resolution, background color, anti-aliasing and margin are all set in a dialog when the script runs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigmaWithDialog.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyAsPngLikeFigmaWithDialog"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigmaWithDialog.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigmaWithDialog.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var uiLang = getCurrentLang();
    var LABELS = {
        dialogTitle: { ja: "ビットマップとしてコピー", en: "Copy as PNG" },
        dpi: { ja: "解像度", en: "Resolution" },
        dpiUnit: { ja: "（dpi）", en: "(dpi)" },
        background: { ja: "背景", en: "Background" },
        transparent: { ja: "透明", en: "Transparent" },
        white: { ja: "白", en: "White" },
        black: { ja: "黒", en: "Black" },
        antialias: { ja: "アンチエイリアスを有効にする", en: "Enable Anti-Aliasing" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" },
        processing: { ja: "処理を開始しています...", en: "Starting process..." },
        done: { ja: "完了しました。", en: "Done." },
        alertNoSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        copiedMsg: { ja: "dpi（", en: "dpi (" },
        copiedMsgSuffix: { ja: "）でビットマップコピーしました。", en: ") bitmap copied." },
        step1: { ja: "一時レイヤー作成中...", en: "Creating temporary layer..." },
        step2: { ja: "複製中...", en: "Duplicating..." },
        step3: { ja: "ラスタライズ範囲作成中...", en: "Creating rasterize bounds..." },
        step4: { ja: "ラスタライズ中...", en: "Rasterizing..." },
        step5: { ja: "拡大中...", en: "Resizing..." },
        step6: { ja: "コピー中...", en: "Copying..." },
        step7: { ja: "一時オブジェクト削除中...", en: "Deleting temporary objects..." },
        margin: { ja: "余白", en: "Margin" },
        tipDpi: { ja: "コピーするビットマップの解像度です。数値が大きいほど精細になり、処理時間も伸びます。", en: "Resolution of the copied bitmap. Higher values are sharper but take longer." },
        tipBackground: { ja: "ビットマップの背景です。「透明」はアルファ付きでコピーします。", en: "Background of the bitmap. Transparent copies it with an alpha channel." },
        tipMargin: { ja: "選択範囲の外側に足す余白です。", en: "Extra space added around the selection." },
        tipAntialias: { ja: "輪郭をなめらかにします。オフにするとドットのままコピーします。", en: "Smooths the edges. Off copies the pixels as they are." }
    };

    function createDialog() {
        var dlg = new Window("dialog", LABELS.dialogTitle[uiLang]);
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        var dpiGroup = dlg.add("group");
        dpiGroup.orientation = "row";
        dpiGroup.add("statictext", undefined, LABELS.dpi[uiLang] + ":");
        var dpiDropdown = dpiGroup.add("dropdownlist", undefined, ["72", "150", "300", "600", "1200"]);
        dpiDropdown.helpTip = LABELS.tipDpi[uiLang];
        dpiGroup.add("statictext", undefined, LABELS.dpiUnit[uiLang]);
        dpiDropdown.selection = 3;

        var bgGroup = dlg.add("group");
        bgGroup.orientation = "row";
        bgGroup.alignChildren = "left";
        bgGroup.add("statictext", undefined, LABELS.background[uiLang] + (uiLang === "ja" ? "：" : ": "));
        var bgTransparent = bgGroup.add("radiobutton", undefined, LABELS.transparent[uiLang]);
        bgTransparent.helpTip = LABELS.tipBackground[uiLang];
        var bgWhite = bgGroup.add("radiobutton", undefined, LABELS.white[uiLang]);
        bgWhite.helpTip = LABELS.tipBackground[uiLang];
        var bgBlack = bgGroup.add("radiobutton", undefined, LABELS.black[uiLang]);
        bgBlack.helpTip = LABELS.tipBackground[uiLang];
        bgWhite.value = true;

        var marginGroup = dlg.add("group");
        marginGroup.orientation = "row";
        marginGroup.add("statictext", undefined, LABELS.margin[uiLang] + ":");
        var marginInput = marginGroup.add("edittext", undefined, "0");
        marginInput.helpTip = LABELS.tipMargin[uiLang];
        marginInput.characters = 4;

        var aaGroup = dlg.add("group");
        var aaCheckbox = aaGroup.add("checkbox", undefined, LABELS.antialias[uiLang]);
        aaCheckbox.helpTip = LABELS.tipAntialias[uiLang];
        aaCheckbox.value = true;

        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "right";
        btnGroup.add("button", undefined, LABELS.cancel[uiLang]);
        btnGroup.add("button", undefined, LABELS.ok[uiLang]);

        return {
            dialog: dlg,
            elements: {
                dpiDropdown: dpiDropdown,
                bgTransparent: bgTransparent,
                bgWhite: bgWhite,
                bgBlack: bgBlack,
                aaCheckbox: aaCheckbox,
                marginInput: marginInput
            }
        };
    }

    function showRasterizeDialog() {
        var ui = createDialog();
        var dlg = ui.dialog;

        if (dlg.show() != 1) return null;

        var e = ui.elements;
        return {
            dpi: parseInt(e.dpiDropdown.selection.text, 10),
            background: e.bgTransparent.value ? "transparent" : (e.bgBlack.value ? "black" : "white"),
            antiAliasing: e.aaCheckbox.value,
            margin: parseInt(e.marginInput.text, 10)
        };
    }

    function runRasterizeWithSettings(settings) {
        var doc = app.activeDocument;
        var originalSelection = app.selection;

        var progressWin = new Window("palette", "処理中...");
        progressWin.pbar = progressWin.add("progressbar", [20, 20, 300, 10], 0, 100);
        progressWin.st = progressWin.add("statictext", undefined, LABELS.processing[uiLang]);
        progressWin.show();

        function updateProgress(value, text) {
            progressWin.pbar.value = value;
            progressWin.st.text = text;
            progressWin.update();
        }

        try {
            updateProgress(10, LABELS.step1[uiLang]);
            var tempLayer = doc.layers.add();
            tempLayer.name = "__TEMP_LAYER__";
            tempLayer.locked = false;
            tempLayer.visible = true;

            updateProgress(20, LABELS.step2[uiLang]);
            var duplicatedItems = [];
            var tempGroup = tempLayer.groupItems.add();
            for (var i = 0; i < originalSelection.length; i++) {
                var dup = originalSelection[i].duplicate(tempGroup, ElementPlacement.PLACEATEND);
                duplicatedItems.push(dup);
            }

            updateProgress(35, LABELS.step3[uiLang]);
            var bounds = tempGroup.visibleBounds;
            var margin = settings.margin || 0;
            var rect = doc.pathItems.rectangle(
                bounds[1] + margin,
                bounds[0] - margin,
                (bounds[2] - bounds[0]) + margin * 2,
                (bounds[1] - bounds[3]) + margin * 2
            );
            rect.stroked = false;
            rect.filled = false;
            rect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);

            updateProgress(50, LABELS.step4[uiLang]);
            var options = new RasterizeOptions();
            options.resolution = settings.dpi;
            options.transparency = (settings.background === "transparent");
            options.antiAliasing = settings.antiAliasing;
            if (!options.transparency) {
                options.backgroundBlack = (settings.background === "black");
            }

            var rasterized = doc.rasterize(tempGroup, rect.geometricBounds, options);

            updateProgress(70, LABELS.step5[uiLang]);
            var resizeRatio = Math.round((settings.dpi / 72) * 100);
            rasterized.resize(resizeRatio, resizeRatio, true, true, true, true, resizeRatio, Transformation.CENTER);

            updateProgress(85, LABELS.step6[uiLang]);
            app.selection = [rasterized];
            app.executeMenuCommand("copy");

            updateProgress(95, LABELS.step7[uiLang]);
            try { rasterized.remove(); } catch (e) {}
            try { rect.remove(); } catch (e) {}
            for (var j = 0; j < duplicatedItems.length; j++) {
                try { duplicatedItems[j].remove(); } catch (e) {}
            }
            try { tempGroup.remove(); } catch (e) {}
            try { tempLayer.remove(); } catch (e) {}

            app.selection = originalSelection;

            updateProgress(100, LABELS.done[uiLang]);

            progressWin.close();
            alert(settings.dpi + LABELS.copiedMsg[uiLang] + resizeRatio + "%" + LABELS.copiedMsgSuffix[uiLang]);
        } catch (err) {
            progressWin.close();
            alert("エラーが発生しました: " + err.message);
        }
    }

    function main() {
        if (app.documents.length === 0 || app.selection.length === 0) {
            alert(LABELS.alertNoSelection[uiLang]);
            return;
        }

        var settings = showRasterizeDialog();
        if (settings) {
            runRasterizeWithSettings(settings);
        }
    }

    main();

})();
