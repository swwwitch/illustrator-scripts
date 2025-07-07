#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

CopyAsPngLikeFigma.jsx

### 概要

- 選択中のオブジェクトを高解像度でラスタライズし、Figmaの「Copy as PNG」のようにビットマップとしてクリップボードへコピーするスクリプトです。
- 600dpiでラスタライズ後、72ppi相当の偶数整数倍率に拡大して最適化します。

### 主な機能

- 一時レイヤー・グループを自動作成して管理
- 600dpiラスタライズ後、拡大倍率を偶数整数に調整
- プログレスバーによる処理進行の可視化
- 一時オブジェクトの自動削除と元選択の復元
- 日本語／英語環境対応

### 処理の流れ

1. オブジェクトを選択
2. 一時レイヤー・グループに複製
3. 600dpiでラスタライズ
4. 72ppi相当の偶数整数倍率に拡大
5. クリップボードにビットマップコピー
6. 一時オブジェクト削除、選択を復元

### 更新履歴

- v1.0.0 (20250502) : 初期バージョン
- v1.0.1 (20250603) : 拡大倍率を偶数整数に調整、コメント整理

---

### Script Name:

CopyAsPngLikeFigma.jsx

### Overview

- A script to rasterize selected objects at high resolution and copy as bitmap to clipboard, similar to Figma’s “Copy as PNG”.
- Rasterizes at 600dpi, then scales up to the nearest even integer multiple of 72ppi for optimization.

### Main Features

- Automatically creates and manages temporary layers and groups
- Rasterizes at 600dpi, adjusts scaling to an even integer multiple
- Progress bar shows processing status
- Automatically deletes temporary objects and restores original selection
- Supports Japanese and English environments

### Process Flow

1. Select objects
2. Duplicate to temporary layer and group
3. Rasterize at 600dpi
4. Scale up to even integer multiple of 72ppi
5. Copy bitmap to clipboard
6. Delete temporary objects and restore selection

### Update History

- v1.0.0 (20250502): Initial version
- v1.0.1 (20250603): Adjusted scaling to even integer multiples, cleaned up comments
*/

// -------------------------------
// 日英ラベル定義 Define label
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();
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
        alert(LABELS.alertSelect[lang]);
        return;
    }

    var doc = app.activeDocument;
    var originalSelection = app.selection;

    // プログレスバー付きダイアログの作成
    var progressWin = new Window("palette", LABELS.progressTitle[lang]);
    progressWin.pbar = progressWin.add("progressbar", [20, 20, 300, 10], 0, 100);
    progressWin.st = progressWin.add("statictext", undefined, LABELS.progressStart[lang]);
    progressWin.show();
    function updateProgress(value, text) {
        progressWin.pbar.value = value;
        progressWin.st.text = text;
        progressWin.update();
    }

    try {
        updateProgress(10, LABELS.progressLayer[lang]);
        // 一時レイヤーの作成
        var tempLayer = doc.layers.add();
        tempLayer.name = "__TEMP_LAYER__";
        tempLayer.locked = false;
        tempLayer.visible = true;

        updateProgress(20, LABELS.progressDuplicate[lang]);
        // オブジェクトを複製して一時グループに追加
        var duplicatedItems = [];
        var tempGroup = tempLayer.groupItems.add();
        for (var i = 0; i < originalSelection.length; i++) {
            var dup = originalSelection[i].duplicate(tempGroup, ElementPlacement.PLACEATEND);
            duplicatedItems.push(dup);
        }

        updateProgress(35, LABELS.progressRect[lang]);
        var bounds = tempGroup.visibleBounds;
        var rect = doc.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
        rect.stroked = false;
        rect.filled = false;
        rect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);

        updateProgress(50, LABELS.progressRaster[lang]);
        var resolution = 600;
        var options = new RasterizeOptions();
        options.resolution = resolution;
        options.transparency = false;
        options.backgroundBlack = false;
        options.antiAliasing = true;

        var rasterized = doc.rasterize(tempGroup, rect.geometricBounds, options);

        updateProgress(70, LABELS.progressResize[lang]);
        // 拡大倍率を「72ppi相当」を基準に偶数の整数に調整
        var baseRatio = (resolution / 72) * 100;
        var resizeRatio = Math.ceil(baseRatio / 2) * 2; // 偶数の整数倍に切り上げ
        rasterized.resize(resizeRatio, resizeRatio);

        updateProgress(85, LABELS.progressCopy[lang]);
        app.selection = [rasterized];
        app.executeMenuCommand("copy");

        updateProgress(95, LABELS.progressClean[lang]);
        try { rasterized.remove(); } catch (e) {}
        try { rect.remove(); } catch (e) {}
        for (var j = 0; j < duplicatedItems.length; j++) {
            try { duplicatedItems[j].remove(); } catch (e) {}
        }
        try { tempGroup.remove(); } catch (e) {}
        try { tempLayer.remove(); } catch (e) {}

        app.selection = originalSelection;

        updateProgress(100, LABELS.progressDone[lang]);
    } catch (err) {
        progressWin.close();
        alert(LABELS.error[lang] + err.message);
        return;
    }

    progressWin.close();
}

main();