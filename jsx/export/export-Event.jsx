#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  ### 概要

  アクティブドキュメントの全アートボードを、名前ごとのルールで PNG 書き出しします。

  ### 書き出しルール

  - "title" / "title2"（"-..." 付きを含む） … 背景「透明」で 100% 書き出し
  - "Doorkeeper" … 背景「白」で 100% と 200% の 2 種類書き出し（200% は "-200" を付加）
  - "シンボル一覧" … 書き出し対象外
  - 上記以外 … 背景「白」で 100% 書き出し

  ルールを追加・変更する場合は `buildExportJobs()` を編集してください。
  空配列を返せば「除外」、複数要素を返せば「複数倍率の書き出し」になります。

  ### 出力

  - 保存先 … ドキュメントと同じフォルダ
  - ファイル名 … "<ドキュメント名>-<アートボード名>[suffix].png"
  - 書き出し後は macOS のみ Finder で保存先を自動オープン

  ### 更新履歴

  - 2025-04-22 … 初版
  - 2026-06-03 … 書き出し前に "Guides Preview for Trim View" レイヤー（"*" 付き含む）を非表示にし、書き出し後に再表示 / 未保存ドキュメントのガードを追加
  - 2026-06-15 … "Guides Preview for Trim View" レイヤー（"*" 付き含む）を非表示ではなく書き出し前に削除するよう変更
  - 2026-06-17 … 定数の巻き上げで削除が機能していなかった不具合を修正。仕様を再び「書き出し前に非表示 → 書き出し後に再表示」に戻す（中断・エラー時も再表示）
  - 2026-07-03 … 進捗ウィンドウを native progressbar に戻す / "Guides Preview for Trim View" レイヤーの非表示・再表示処理を削除
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v1.0.4";

    // =========================================
    // メイン処理 / Main routine
    // =========================================

    exportArtboardsAsPng();

    /* 全アートボードを名前ごとのルールで PNG 書き出し / Export every artboard as PNG using name-based rules */
    function exportArtboardsAsPng() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDoc = app.activeDocument;
        /* 一度も保存していないドキュメントは保存先が定まらないため中止 / Abort when the document has never been saved (no output folder) */
        var outputFolder = null;
        try {
            outputFolder = activeDoc.fullName.parent;
        } catch (e) {}
        if (!outputFolder || !outputFolder.exists) {
            alert("先にドキュメントを保存してください。");
            return;
        }

        /* 書き出し対象とジョブを一括算出（buildExportJobs の二度呼びを回避）/ Resolve targets and jobs in one pass */
        var plan = buildExportPlan(activeDoc);
        if (plan.totalJobs === 0) {
            return;
        }

        var baseFileName = activeDoc.name.replace(/\.ai$/i, "");
        var progress = createProgressWindow(plan.totalJobs);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var cancelled = false;
        try {
            var completed = 0;
            for (var i = 0; i < plan.items.length && !cancelled; i++) {
                var item = plan.items[i];
                activeDoc.artboards.setActiveArtboardIndex(item.index);
                for (var j = 0; j < item.jobs.length; j++) {
                    progress.update(completed, item.name + item.jobs[j].suffix);
                    /* キャンセルボタンが押されていれば中断 / Stop if the cancel button was pressed */
                    if (progress.isCancelled()) {
                        cancelled = true;
                        break;
                    }
                    exportArtboardAsPng(activeDoc, outputFolder, baseFileName, item.name, item.jobs[j]);
                    completed++;
                }
            }
            progress.update(completed, cancelled ? "キャンセルしました / Cancelled" : "完了 / Done");
        } finally {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            progress.close();
        }

        /* キャンセル時は Finder を開かない / Skip revealing the folder when cancelled */
        if (!cancelled) {
            revealOutputFolder(outputFolder);
        }
    }

    // =========================================
    // ルール判定 / Rule resolver
    // =========================================

    /* 書き出し対象のアートボードとジョブ・総数を一括算出 / Resolve target artboards, their jobs, and the total count */
    function buildExportPlan(doc) {
        var items = [];
        var totalJobs = 0;
        var artboardCount = doc.artboards.length;
        for (var i = 0; i < artboardCount; i++) {
            var artboardName = doc.artboards[i].name;
            var jobs = buildExportJobs(artboardName);
            if (jobs.length === 0) {
                continue;
            }
            items.push({ index: i, name: artboardName, jobs: jobs });
            totalJobs += jobs.length;
        }
        return { items: items, totalJobs: totalJobs };
    }

    /* アートボード名から書き出しジョブ配列を生成 / Build export jobs from an artboard name (empty array = skip) */
    function buildExportJobs(artboardName) {
        /* シンボル一覧は除外 / Skip "シンボル一覧" */
        if (artboardName === "シンボル一覧") {
            return [];
        }
        /* Doorkeeper は 200% + 100% を白背景で / Doorkeeper: 200% and 100% with white background */
        if (artboardName === "Doorkeeper") {
            return [
                { scale: 200, transparent: false, suffix: "-200" },
                { scale: 100, transparent: false, suffix: "" }
            ];
        }
        /* title / title2 系は透明背景で 100% / title / title2 family: 100% with transparent background */
        if (/^(title|title2)(-|$)/.test(artboardName)) {
            return [{ scale: 100, transparent: true, suffix: "" }];
        }
        /* その他は白背景で 100% / Otherwise: 100% with white background */
        return [{ scale: 100, transparent: false, suffix: "" }];
    }

    // =========================================
    // 書き出しヘルパー / Export helper
    // =========================================

    /* 1 アートボードを指定倍率・背景で PNG 書き出し / Export a single artboard as PNG at the given scale and background */
    function exportArtboardAsPng(sourceDoc, outputFolder, baseFileName, artboardName, job) {
        var exportOptions = new ExportOptionsPNG24();
        exportOptions.artBoardClipping = true;
        exportOptions.antiAliasing = true;
        exportOptions.transparency = job.transparent;
        exportOptions.horizontalScale = job.scale;
        exportOptions.verticalScale = job.scale;

        var outputFileName = baseFileName + "-" + artboardName + job.suffix + ".png";
        var outputFile = new File(outputFolder.fsName + "/" + outputFileName);

        try {
            sourceDoc.exportFile(outputFile, ExportType.PNG24, exportOptions);
        } catch (e) {
            alert("アートボード「" + artboardName + "」の書き出し中にエラーが発生しました：\n" + e.message);
        }
    }

    /* macOS のみ保存先を Finder で開く / Reveal the output folder in Finder (macOS only) */
    function revealOutputFolder(folder) {
        if (Folder.fs === "Macintosh") {
            folder.execute();
        }
    }

    // =========================================
    // 進捗表示 / Progress UI
    // =========================================

    /* 進捗ウィンドウを生成 / Build a palette window with a status label, a native progress bar, and a cancel button */
    function createProgressWindow(total) {
        var win = new Window("palette", "PNG 書き出し中… " + SCRIPT_VERSION, undefined, { closeButton: false });
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = [16, 16, 16, 16];
        win.spacing = 10;

        var statusText = win.add("statictext", undefined, "準備中… / Preparing…");
        statusText.preferredSize.width = 360;

        var bar = win.add("progressbar", undefined, 0, total);
        bar.preferredSize = [360, 14];

        /* キャンセルボタン（押下でフラグを立て、ループ側が中断）/ Cancel button (sets a flag that the export loop checks) */
        var cancelled = false;
        var buttonGroup = win.add("group");
        buttonGroup.alignment = ["right", "top"];
        var cancelButton = buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        cancelButton.onClick = function () {
            cancelled = true;
            cancelButton.enabled = false;
            statusText.text = "キャンセル中…";
            win.update();
        };

        win.show();

        return {
            isCancelled: function () {
                return cancelled;
            },
            update: function (value, label) {
                statusText.text = label + "  (" + value + " / " + total + ")";
                bar.value = value;
                win.update();
            },
            close: function () {
                win.close();
            }
        };
    }
})();