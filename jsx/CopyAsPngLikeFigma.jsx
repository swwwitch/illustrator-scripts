#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * スクリプトの概要：
 * 選択中のオブジェクトを一時的にレイヤーに複製し、600dpiでラスタライズ。
 * ラスタライズ後、72ppi相当（833%）のサイズに最も近い「偶数の整数倍」に拡大し、クリップボードへビットマップとしてコピー。
 * プログレスバー付きUIにより処理の進行を可視化し、一時レイヤーや一時オブジェクトは後に自動削除。
 *
 * 作成日：2025-05-02
 * 最終更新日：2025-06-03（拡大倍率を偶数の整数倍に調整し、コメント・構成を整理）
 */

function main() {
    if (app.documents.length === 0 || app.selection.length === 0) {
        alert("オブジェクトを選択してください。");
        return;
    }

    var doc = app.activeDocument;
    var originalSelection = app.selection;

    // プログレスバー付きダイアログの作成
    var progressWin = new Window("palette", "処理中...");
    progressWin.pbar = progressWin.add("progressbar", [20, 20, 300, 10], 0, 100);
    progressWin.st = progressWin.add("statictext", undefined, "処理を開始しています...");
    progressWin.show();
    function updateProgress(value, text) {
        progressWin.pbar.value = value;
        progressWin.st.text = text;
        progressWin.update();
    }

    try {
        updateProgress(10, "一時レイヤー作成中...");
        // 一時レイヤーの作成
        var tempLayer = doc.layers.add();
        tempLayer.name = "__TEMP_LAYER__";
        tempLayer.locked = false;
        tempLayer.visible = true;

        updateProgress(20, "複製中...");
        // オブジェクトを複製して一時グループに追加
        var duplicatedItems = [];
        var tempGroup = tempLayer.groupItems.add();
        for (var i = 0; i < originalSelection.length; i++) {
            var dup = originalSelection[i].duplicate(tempGroup, ElementPlacement.PLACEATEND);
            duplicatedItems.push(dup);
        }

        updateProgress(35, "ラスタライズ範囲作成中...");
        var bounds = tempGroup.visibleBounds;
        var rect = doc.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
        rect.stroked = false;
        rect.filled = false;
        rect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);

        updateProgress(50, "ラスタライズ中...");
        var resolution = 600;
        var options = new RasterizeOptions();
        options.resolution = resolution;
        options.transparency = false;
        options.backgroundBlack = false;
        options.antiAliasing = true;

        var rasterized = doc.rasterize(tempGroup, rect.geometricBounds, options);

        updateProgress(70, "拡大中...");
        // 拡大倍率を「72ppi相当」を基準に偶数の整数に調整
        var baseRatio = (resolution / 72) * 100;
        var resizeRatio = Math.ceil(baseRatio / 2) * 2; // 偶数の整数倍に切り上げ
        rasterized.resize(resizeRatio, resizeRatio);

        updateProgress(85, "コピー中...");
        app.selection = [rasterized];
        app.executeMenuCommand("copy");

        updateProgress(95, "一時オブジェクト削除中...");
        try { rasterized.remove(); } catch (e) {}
        try { rect.remove(); } catch (e) {}
        for (var j = 0; j < duplicatedItems.length; j++) {
            try { duplicatedItems[j].remove(); } catch (e) {}
        }
        try { tempGroup.remove(); } catch (e) {}
        try { tempLayer.remove(); } catch (e) {}

        app.selection = originalSelection;

        updateProgress(100, "完了しました。");
    } catch (err) {
        progressWin.close();
        alert("エラーが発生しました: " + err.message);
        return;
    }

    progressWin.close();
}

main();