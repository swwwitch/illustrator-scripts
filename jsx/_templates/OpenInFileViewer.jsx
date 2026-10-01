#target illustrator

/*

### 概要

フォルダーを開く・ファイルを選択して表示する処理を、Path Finder が起動していれば Path Finder で、そうでなければ Finder で行う再利用テンプレートです。
補助アプリ /Applications/OpenInFileViewer.app が無い環境では false を返すので、呼び出し側で execute() に落とします。

### Overview

A reusable template that opens a folder or reveals a file in Path Finder when it is running, otherwise in Finder.
Without the helper app /Applications/OpenInFileViewer.app it returns false, so the caller falls back to execute().

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "OpenInFileViewer";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る。識別子は openInFileViewer
    // 2. 補助アプリは illustrator-scripts の helpers/OpenInFileViewer.applescript から作る:
    //      osacompile -o /Applications/OpenInFileViewer.app helpers/OpenInFileViewer.applescript
    // 3. 補助アプリが無いときは false が返るので、これまでの処理に落とす:
    //      if (!openInFileViewer(outputFolder)) outputFolder.execute();
    //      if (!openInFileViewer(exportedFile)) exportedFile.parent.execute();
    // 4. フォルダーは開き、ファイルは選択して表示する。Path Finder でファイルを選択表示するときだけ、初回にオートメーションの許可を求められる

    // ファイルビューアで表示（再利用パーツ） / Show in file viewer (reusable)

    /**
     * Path Finder（起動中のとき）か Finder で、フォルダーを開くかファイルを選択して表示する。
     * 補助アプリ /Applications/OpenInFileViewer.app に一時ファイルでパスを渡して起動する。
     * 補助アプリは illustrator-scripts の helpers/OpenInFileViewer.applescript から作る
     * @param {File|Folder} targetItem - 開くフォルダーか、選択して表示するファイル
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function openInFileViewer(targetItem) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var viewerAppPath = "/Applications/OpenInFileViewer.app";
        var pathFilePath = "/tmp/open_in_file_viewer_path.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(viewerAppPath).exists && !new File(viewerAppPath).exists) return false;

        var pathFile = new File(pathFilePath);
        var written = false;
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (pathFile.open("w")) {
                /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
                written = pathFile.write(targetItem.fsName);
            }
        } catch (e) {
        } finally {
            try { pathFile.close(); } catch (closeError) {}
        }
        return written && new File(viewerAppPath).execute();
    }

    // ファイルビューアで表示（再利用パーツ）ここまで / End of the reusable file viewer

    // =========================================
    // 動作確認 / Demo
    // =========================================
    var demoFolder = Folder.selectDialog();
    if (demoFolder && !openInFileViewer(demoFolder)) demoFolder.execute();

})();
