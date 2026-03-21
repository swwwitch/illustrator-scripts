#target illustrator

(function () {
    // 選択オブジェクトの確認
    if (app.selection.length === 0) {
        alert("テキストを選択してください。");
        return;
    }

    var sel = app.selection[0];

    // テキストフレームか確認
    if (sel.typename !== "TextFrame") {
        alert("選択されたオブジェクトはテキストではありません。");
        return;
    }

    // 元のテキストを取得
    var originalContents = sel.contents;

    // 改行・スペースを除去して文字だけ取り出す（必要に応じて調整）
    // var chars = originalContents.replace(/[\r\n]/g, "");

    // 1文字ずつ \r で結合
    var newContents = originalContents.split("").join("\r");

    // テキストを書き換え
    sel.contents = newContents;

    // テキストフレームをエリアテキストにしている場合は幅を調整
    // （ポイントテキストならそのまま縦に伸びる）

})();
