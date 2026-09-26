#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorに登録されているアクションセットを、デスクトップへ書き出します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportActions.md

### Overview

Exports the action sets registered in Illustrator to the desktop.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportActions.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportActions";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportActions.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportActions.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 書き出し先（デスクトップ直下のフォルダー名）/ Destination folder name on the desktop */
    var EXPORT_FOLDER_NAME = "Illustrator_Actions";

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
            noActionSetsApi: {
                ja: "このバージョンのIllustratorでは actionSets にアクセスできません。スクリプトではアクションセット一覧を取得できない仕様です。",
                en: "This version of Illustrator does not expose actionSets, so scripts cannot list the action sets."
            },
            noActionSets: { ja: "アクションセットが見つかりません。", en: "No action sets were found." }
        },
        report: {
            title:     { ja: "=== 書き出し完了 ===", en: "=== Export complete ===" },
            folder:    { ja: "保存先: ", en: "Saved to: " },
            succeeded: { ja: "✅ 成功したアクションセット (%1件):", en: "✅ Exported action sets (%1):" },
            failed:    { ja: "❌ 失敗したアクションセット (%1件):", en: "❌ Failed action sets (%1):" },
            errorNote: { ja: "（エラー: %1）", en: " (error: %1)" }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "alert.noActionSets" のようなドット区切りのキー
     * @param {string|number} [replacement] - ラベル中の %1 に入れる値
     * @returns {string} 現在のUI言語のラベル
     */
    function getLabel(labelPath, replacement) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var label = (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
        return (replacement === undefined) ? label : label.split("%1").join(String(replacement));
    }

    // =========================================
    // 書き出し / Export
    // =========================================

    /**
     * アクションセット名をファイル名に使える形にする
     * @param {string} actionSetName - アクションセット名
     * @returns {string} ファイル名に使えない文字を「_」に置き換えた名前
     */
    function toSafeFileName(actionSetName) {
        return actionSetName.replace(new RegExp('[\\\\/:*?"<>|]', 'g'), "_");
    }

    /**
     * 名前の一覧を「  ・名前」の行にする
     * @param {string[]} names - 並べる名前
     * @returns {string} 1行ずつ改行で終わる文字列
     */
    function formatNameLines(names) {
        var lines = "";
        for (var i = 0; i < names.length; i++) {
            lines += "  ・" + names[i] + "\n";
        }
        return lines;
    }

    /**
     * 書き出し結果の報告文を組み立てる
     * @param {Folder} saveFolder - 保存先フォルダー
     * @param {string[]} exportedList - 書き出せたアクションセット名
     * @param {string[]} errorList - 書き出せなかったアクションセット名（エラー内容付き）
     * @returns {string} 報告文
     */
    function buildReportMessage(saveFolder, exportedList, errorList) {
        var reportMessage = getLabel("report.title") + "\n";
        reportMessage += getLabel("report.folder") + saveFolder.fsName + "\n\n";

        if (exportedList.length > 0) {
            reportMessage += getLabel("report.succeeded", exportedList.length) + "\n";
            reportMessage += formatNameLines(exportedList);
        }

        if (errorList.length > 0) {
            reportMessage += "\n" + getLabel("report.failed", errorList.length) + "\n";
            reportMessage += formatNameLines(errorList);
        }
        return reportMessage;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 登録されているアクションセットをデスクトップの書き出し先フォルダーへ書き出す
     * @returns {void}
     */
    function main() {
        var actionSets = app.actionSets;
        if (!actionSets) {
            alert(getLabel("alert.noActionSetsApi"));
            return;
        }
        var actionCount = actionSets.length;

        if (actionCount === 0) {
            alert(getLabel("alert.noActionSets"));
            return;
        }

        /* 保存先フォルダーを用意 / Make sure the destination folder exists */
        var saveFolder = new Folder(Folder.desktop + "/" + EXPORT_FOLDER_NAME);
        if (!saveFolder.exists) {
            saveFolder.create();
        }

        var exportedList = [];
        var errorList = [];

        /* 全アクションセットを書き出し / Export every action set */
        for (var i = 0; i < actionCount; i++) {
            var actionSetName = actionSets[i].name;
            var saveFile = new File(saveFolder + "/" + toSafeFileName(actionSetName) + ".aia");

            /* 書き出しの失敗は一覧に記録して続行 / Record a failed save and move on */
            try {
                actionSets[i].save(saveFile);
                exportedList.push(actionSetName);
            } catch (e) {
                errorList.push(actionSetName + getLabel("report.errorNote", e.message));
            }
        }

        alert(buildReportMessage(saveFolder, exportedList, errorList));
    }

    main();

})();
