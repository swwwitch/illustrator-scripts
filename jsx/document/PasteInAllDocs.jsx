#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
------------------------------------------------------------
複数ドキュメントへコピー内容をペースト
------------------------------------------------------------
概要:
- すでにコピー済みのオブジェクトを、すべての開いているドキュメントに
  「同じ位置」でペーストする
- 「pasteInPlace」を使用するため、各ドキュメントの座標にあわせて配置

対象:
- 開いているすべてのドキュメント
- コピー済みのオブジェクト（事前にコピーしておく必要あり）

非対象:
- ペースト先で選択済みのオブジェクトは上書きされない
- コピーしていない場合は警告を表示して終了
------------------------------------------------------------
*/

function main() {
    try {
        // ドキュメントが1つも開いていない場合は終了
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。\rNo documents are open.");
            return;
        }

        // 最初のドキュメントを保持（コピー元とする）
        var sourceDoc = app.activeDocument;

        // コピーされていないとペーストに失敗するのでチェック
        // Illustratorではクリップボードの空判定が難しいため try-catch で対応
        try {
            sourceDoc.activate();
            app.executeMenuCommand("pasteInPlace");
            // ペーストされたものを削除してクリップボードを維持
            app.cut();
        } catch (e) {
            alert("コピーされたオブジェクトがありません。\rNo objects copied.");
            return;
        }

        // 各ドキュメントへペースト
        for (var i = 0; i < app.documents.length; i++) {
            var doc = app.documents[i];
            doc.activate();
            try {
                app.executeMenuCommand("pasteInPlace");
            } catch (e) {
                alert("ペーストに失敗しました: " + doc.name + "\rFailed to paste in: " + doc.name);
            }
        }

        // 元のドキュメントを再度アクティブに
        sourceDoc.activate();
        alert("コピーしたオブジェクトをすべてのドキュメントにペーストしました。\rObjects pasted into all documents.");

    } catch (err) {
        alert("エラーが発生しました:\r" + err);
    }
}

main();