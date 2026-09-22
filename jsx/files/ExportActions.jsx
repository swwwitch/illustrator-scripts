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
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportActions.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportActions.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
        var reportMessage = "=== 書き出し完了 ===\n";
        reportMessage += "保存先: " + saveFolder.fsName + "\n\n";

        if (exportedList.length > 0) {
            reportMessage += "✅ 成功したアクションセット (" + exportedList.length + "件):\n";
            reportMessage += formatNameLines(exportedList);
        }

        if (errorList.length > 0) {
            reportMessage += "\n❌ 失敗したアクションセット (" + errorList.length + "件):\n";
            reportMessage += formatNameLines(errorList);
        }
        return reportMessage;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 登録されているアクションセットをデスクトップの Illustrator_Actions フォルダーへ書き出す
     * @returns {void}
     */
    function main() {
        var actionSets = app.actionSets ? app.actionSets : null;
        if (!actionSets) {
            alert("このバージョンのIllustratorでは actionSets にアクセスできません。スクリプトではアクションセット一覧を取得できない仕様です。");
            return;
        }
        var actionCount = actionSets.length;

        if (actionCount === 0) {
            alert("アクションセットが見つかりません。");
            return;
        }

        /* 保存先フォルダーを用意 / Make sure the destination folder exists */
        var saveFolder = new Folder(Folder.desktop + "/Illustrator_Actions");
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
                errorList.push(actionSetName + "（エラー: " + e.message + "）");
            }
        }

        alert(buildReportMessage(saveFolder, exportedList, errorList));
    }

    main();

})();
