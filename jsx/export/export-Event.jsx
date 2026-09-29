#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブドキュメントのすべてのアートボードを、アートボード名ごとのルールでPNG書き出しします。
背景の透明・白、倍率、書き出し対象外の判定は `buildExportJobs()` で定義します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md

### Overview

Exports every artboard of the active document to PNG, using rules keyed on the artboard name.
Transparent or white background, scale, and exclusions are all defined in `buildExportJobs()`.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export-Event.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "export-Event";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export-Event.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * アートボード名から書き出しジョブの配列を作る（空配列は書き出し対象外）
     * @param {string} artboardName - アートボード名
     * @returns {Array<{scale: number, transparent: boolean, suffix: string}>} 書き出しジョブ
     */
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
    // レイアウト / Layout
    // =========================================
    var PROGRESS_MARGINS    = [16, 16, 16, 16]; /* 進捗ウィンドウの余白 [左,上,右,下] / Progress window margins */
    var PROGRESS_SPACING    = 10;               /* 進捗ウィンドウ内の間隔 / Progress window spacing */
    var PROGRESS_WIDTH      = 360;              /* 状況表示とバーの幅 / Width of the status text and bar */
    var PROGRESS_BAR_HEIGHT = 14;               /* バーの高さ / Bar height */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            progressTitle: { ja: "PNG 書き出し中…", en: "Exporting PNG…" }
        },
        status: {
            preparing:  { ja: "準備中…", en: "Preparing…" },
            cancelling: { ja: "キャンセル中…", en: "Cancelling…" },
            cancelled:  { ja: "キャンセルしました", en: "Cancelled" },
            done:       { ja: "完了", en: "Done" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            saveFirst: { ja: "先にドキュメントを保存してください。", en: "Save the document first." },
            exportError: {
                ja: "アートボード「%1」の書き出し中にエラーが発生しました：\n",
                en: "An error occurred while exporting the artboard \"%1\":\n"
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 全アートボードを名前ごとのルールで PNG 書き出しする
     * @returns {void}
     */
    function exportArtboardsAsPng() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDoc = app.activeDocument;
        /* 一度も保存していないドキュメントは保存先が定まらないため中止（fullName が例外になることがある）/ Abort when the document has never been saved (fullName may throw) */
        var outputFolder = null;
        try {
            outputFolder = activeDoc.fullName.parent;
        } catch (e) {}
        if (!outputFolder || !outputFolder.exists) {
            alert(getLabel("alert.saveFirst"));
            return;
        }

        /* 書き出し対象とジョブを一括算出（buildExportJobs の二度呼びを回避）/ Resolve targets and jobs in one pass */
        var exportPlan = buildExportPlan(activeDoc);
        if (exportPlan.totalJobs === 0) {
            return;
        }

        var baseFileName = activeDoc.name.replace(/\.ai$/i, "");
        var progress = createProgressWindow(exportPlan.totalJobs);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var cancelled = false;
        /* 途中で失敗しても警告表示の設定と進捗ウィンドウは戻す / Always restore alerts and close the progress window */
        try {
            var completedCount = 0;
            for (var i = 0; i < exportPlan.artboardPlans.length && !cancelled; i++) {
                var artboardPlan = exportPlan.artboardPlans[i];
                activeDoc.artboards.setActiveArtboardIndex(artboardPlan.index);
                for (var j = 0; j < artboardPlan.exportJobs.length; j++) {
                    progress.update(completedCount, artboardPlan.name + artboardPlan.exportJobs[j].suffix);
                    /* キャンセルボタンが押されていれば中断 / Stop if the cancel button was pressed */
                    if (progress.isCancelled()) {
                        cancelled = true;
                        break;
                    }
                    exportArtboardAsPng(activeDoc, outputFolder, baseFileName, artboardPlan.name, artboardPlan.exportJobs[j]);
                    completedCount++;
                }
            }
            progress.update(completedCount, getLabel(cancelled ? "status.cancelled" : "status.done"));
        } finally {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            progress.close();
        }
    }

    // =========================================
    // ルール判定 / Rule resolver
    // =========================================

    /**
     * 書き出し対象のアートボードとジョブ、ジョブの総数を一括で求める
     * @param {Document} doc - 対象ドキュメント
     * @returns {{artboardPlans: Array<{index: number, name: string, exportJobs: Object[]}>, totalJobs: number}} 書き出し計画
     */
    function buildExportPlan(doc) {
        var artboardPlans = [];
        var totalJobs = 0;
        var artboardCount = doc.artboards.length;
        for (var i = 0; i < artboardCount; i++) {
            var artboardName = doc.artboards[i].name;
            var exportJobs = buildExportJobs(artboardName);
            if (exportJobs.length === 0) {
                continue;
            }
            artboardPlans.push({ index: i, name: artboardName, exportJobs: exportJobs });
            totalJobs += exportJobs.length;
        }
        return { artboardPlans: artboardPlans, totalJobs: totalJobs };
    }

    // =========================================
    // 書き出しヘルパー / Export helper
    // =========================================

    /**
     * 1つのアートボードを指定倍率・背景で PNG 書き出しする（アクティブなアートボードが対象）
     * @param {Document} sourceDoc - 書き出すドキュメント
     * @param {Folder} outputFolder - 保存先フォルダー
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {string} artboardName - アートボード名
     * @param {{scale: number, transparent: boolean, suffix: string}} exportJob - 書き出しジョブ
     * @returns {void}
     */
    function exportArtboardAsPng(sourceDoc, outputFolder, baseFileName, artboardName, exportJob) {
        var exportOptions = new ExportOptionsPNG24();
        exportOptions.artBoardClipping = true;
        exportOptions.antiAliasing = true;
        exportOptions.transparency = exportJob.transparent;
        exportOptions.horizontalScale = exportJob.scale;
        exportOptions.verticalScale = exportJob.scale;

        var outputFileName = baseFileName + "-" + artboardName + exportJob.suffix + ".png";
        var outputFile = new File(outputFolder.fsName + "/" + outputFileName);

        /* 1枚の失敗で全体を止めない / One failed export does not stop the rest */
        try {
            sourceDoc.exportFile(outputFile, ExportType.PNG24, exportOptions);
        } catch (e) {
            alert(getLabel("alert.exportError").split("%1").join(artboardName) + e.message);
        }
    }

    // =========================================
    // 進捗表示 / Progress UI
    // =========================================

    /**
     * 状況表示・プログレスバー・キャンセルボタンを持つ進捗パレットを開く
     * @param {number} totalJobs - ジョブの総数
     * @returns {{isCancelled: Function, update: Function, close: Function}} 進捗パレットの操作
     */
    function createProgressWindow(totalJobs) {
        var progressWin = new Window("palette", getLabel("dialog.progressTitle") + " " + SCRIPT_VERSION, undefined, { closeButton: false });
        progressWin.orientation = "column";
        progressWin.alignChildren = "fill";
        progressWin.margins = PROGRESS_MARGINS;
        progressWin.spacing = PROGRESS_SPACING;

        var statusText = progressWin.add("statictext", undefined, getLabel("status.preparing"));
        statusText.preferredSize.width = PROGRESS_WIDTH;

        var progressBar = progressWin.add("progressbar", undefined, 0, totalJobs);
        progressBar.preferredSize = [PROGRESS_WIDTH, PROGRESS_BAR_HEIGHT];

        /* キャンセルボタン（押下でフラグを立て、ループ側が中断）/ Cancel button (sets a flag that the export loop checks) */
        var cancelled = false;
        var btnRowGroup = progressWin.add("group");
        btnRowGroup.alignment = ["right", "top"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.onClick = function () {
            cancelled = true;
            btnCancel.enabled = false;
            statusText.text = getLabel("status.cancelling");
            progressWin.update();
        };

        progressWin.show();

        return {
            isCancelled: function () {
                return cancelled;
            },
            update: function (value, statusLabel) {
                statusText.text = statusLabel + "  (" + value + " / " + totalJobs + ")";
                progressBar.value = value;
                progressWin.update();
            },
            close: function () {
                progressWin.close();
            }
        };
    }

    exportArtboardsAsPng();

})();
