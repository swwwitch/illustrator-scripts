#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorに登録されているアクションセットを、デスクトップへ書き出します。

詳細は README を参照してください。

### Overview

Exports the action sets registered in Illustrator to the desktop.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportActions";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportActions.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportActions.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    #target illustrator

    (function () {
        // デスクトップのパスを取得
        var desktop = Folder.desktop;

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

        // 保存先フォルダの確認
        var saveFolder = new Folder(desktop + "/Illustrator_Actions");
        if (!saveFolder.exists) {
            saveFolder.create();
        }

        var exportedList = [];
        var errorList = [];

        // 全アクションセットをループして書き出し
        for (var i = 0; i < actionCount; i++) {
            var setName = actionSets[i].name;

            // ファイル名に使えない文字を置換
            var safeFileName = setName.replace(new RegExp('[\\\\/:*?"<>|]', 'g'), "_");
            var saveFile = new File(saveFolder + "/" + safeFileName + ".aia");

            try {
                app.exportPDFPreset; // 疎通確認（任意）
                actionSets[i].save(saveFile);
                exportedList.push(setName);
            } catch (e) {
                errorList.push(setName + "（エラー: " + e.message + "）");
            }
        }

        // 結果レポート
        var msg = "=== 書き出し完了 ===\n";
        msg += "保存先: " + saveFolder.fsName + "\n\n";

        if (exportedList.length > 0) {
            msg += "✅ 成功したアクションセット (" + exportedList.length + "件):\n";
            for (var j = 0; j < exportedList.length; j++) {
                msg += "  ・" + exportedList[j] + "\n";
            }
        }

        if (errorList.length > 0) {
            msg += "\n❌ 失敗したアクションセット (" + errorList.length + "件):\n";
            for (var k = 0; k < errorList.length; k++) {
                msg += "  ・" + errorList[k] + "\n";
            }
        }

        alert(msg);

    })();

})();
