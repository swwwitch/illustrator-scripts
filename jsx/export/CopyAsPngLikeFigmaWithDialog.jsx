#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

CopyAsPngLikeFigmaWithDialog.jsx

### 概要

- 選択オブジェクトを高解像度でラスタライズし、PNG相当のビットマップとしてクリップボードにコピーするIllustrator用スクリプトです。
- 実行時にダイアログで解像度、背景色、アンチエイリアス、余白を設定できます。

### 主な機能

- 解像度（dpi）を選択可能（72〜1200）
- 背景色（透明・白・黒）を指定
- アンチエイリアスON/OFF切替
- 余白の設定
- 日本語／英語インターフェース対応

### 処理の流れ

1. 対象オブジェクトを選択
2. ダイアログで各種設定を指定
3. 一時レイヤーに複製し、ラスタライズと拡大を実行
4. クリップボードにコピー
5. 一時オブジェクトを削除し、選択を復元

### 更新履歴

- v1.0.0 (20250502) : 初期バージョン
- v1.0.1 (20250603) : ラベル整理と拡大率補正

---

### Script Name:

CopyAsPngLikeFigmaWithDialog.jsx

### Overview

- An Illustrator script that rasterizes selected objects at high resolution and copies them as a PNG-like bitmap to the clipboard.
- Allows configuring resolution, background color, anti-aliasing, and margin via a dialog at runtime.

### Main Features

- Selectable resolution (dpi) from 72 to 1200
- Choose background color (transparent, white, black)
- Toggle anti-aliasing on/off
- Set margin
- Japanese and English UI support

### Process Flow

1. Select target objects
2. Configure settings via dialog
3. Duplicate to temporary layer, rasterize and resize
4. Copy to clipboard
5. Delete temporary objects and restore selection

### Update History

- v1.0.0 (20250502): Initial version
- v1.0.1 (20250603): Refined labels and scaling adjustments
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "ビットマップとしてコピー", en: "Copy as PNG" },
    dpi: { ja: "解像度", en: "Resolution" },
    dpiUnit: { ja: "（dpi）", en: "(dpi)" },
    background: { ja: "背景：", en: "Background:" },
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
    margin: { ja: "余白", en: "Margin" }
};

function createDialog() {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var dpiGroup = dlg.add("group");
    dpiGroup.orientation = "row";
    dpiGroup.add("statictext", undefined, LABELS.dpi[lang] + ":");
    var dpiDropdown = dpiGroup.add("dropdownlist", undefined, ["72", "150", "300", "600", "1200"]);
    dpiGroup.add("statictext", undefined, LABELS.dpiUnit[lang]);
    dpiDropdown.selection = 3;

    var bgGroup = dlg.add("group");
    bgGroup.orientation = "row";
    bgGroup.alignChildren = "left";
    bgGroup.add("statictext", undefined, LABELS.background[lang]);
    var bgTransparent = bgGroup.add("radiobutton", undefined, LABELS.transparent[lang]);
    var bgWhite = bgGroup.add("radiobutton", undefined, LABELS.white[lang]);
    var bgBlack = bgGroup.add("radiobutton", undefined, LABELS.black[lang]);
    bgWhite.value = true;

    var marginGroup = dlg.add("group");
    marginGroup.orientation = "row";
    marginGroup.add("statictext", undefined, LABELS.margin[lang] + ":");
    var marginInput = marginGroup.add("edittext", undefined, "0");
    marginInput.characters = 4;

    var aaGroup = dlg.add("group");
    var aaCheckbox = aaGroup.add("checkbox", undefined, LABELS.antialias[lang]);
    aaCheckbox.value = true;

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";
    btnGroup.add("button", undefined, LABELS.cancel[lang]);
    btnGroup.add("button", undefined, LABELS.ok[lang]);

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
    progressWin.st = progressWin.add("statictext", undefined, LABELS.processing[lang]);
    progressWin.show();

    function updateProgress(value, text) {
        progressWin.pbar.value = value;
        progressWin.st.text = text;
        progressWin.update();
    }

    try {
        updateProgress(10, LABELS.step1[lang]);
        var tempLayer = doc.layers.add();
        tempLayer.name = "__TEMP_LAYER__";
        tempLayer.locked = false;
        tempLayer.visible = true;

        updateProgress(20, LABELS.step2[lang]);
        var duplicatedItems = [];
        var tempGroup = tempLayer.groupItems.add();
        for (var i = 0; i < originalSelection.length; i++) {
            var dup = originalSelection[i].duplicate(tempGroup, ElementPlacement.PLACEATEND);
            duplicatedItems.push(dup);
        }

        updateProgress(35, LABELS.step3[lang]);
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

        updateProgress(50, LABELS.step4[lang]);
        var options = new RasterizeOptions();
        options.resolution = settings.dpi;
        options.transparency = (settings.background === "transparent");
        options.antiAliasing = settings.antiAliasing;
        if (!options.transparency) {
            options.backgroundBlack = (settings.background === "black");
        }

        var rasterized = doc.rasterize(tempGroup, rect.geometricBounds, options);

        updateProgress(70, LABELS.step5[lang]);
        var resizeRatio = Math.round((settings.dpi / 72) * 100);
        rasterized.resize(resizeRatio, resizeRatio, true, true, true, true, resizeRatio, Transformation.CENTER);

        updateProgress(85, LABELS.step6[lang]);
        app.selection = [rasterized];
        app.executeMenuCommand("copy");

        updateProgress(95, LABELS.step7[lang]);
        try { rasterized.remove(); } catch (e) {}
        try { rect.remove(); } catch (e) {}
        for (var j = 0; j < duplicatedItems.length; j++) {
            try { duplicatedItems[j].remove(); } catch (e) {}
        }
        try { tempGroup.remove(); } catch (e) {}
        try { tempLayer.remove(); } catch (e) {}

        app.selection = originalSelection;

        updateProgress(100, LABELS.done[lang]);

        progressWin.close();
        alert(settings.dpi + LABELS.copiedMsg[lang] + resizeRatio + "%" + LABELS.copiedMsgSuffix[lang]);
    } catch (err) {
        progressWin.close();
        alert("エラーが発生しました: " + err.message);
    }
}

function main() {
    if (app.documents.length === 0 || app.selection.length === 0) {
        alert(LABELS.alertNoSelection[lang]);
        return;
    }

    var settings = showRasterizeDialog();
    if (settings) {
        runRasterizeWithSettings(settings);
    }
}

main();