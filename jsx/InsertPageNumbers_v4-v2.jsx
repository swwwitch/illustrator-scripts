#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    スクリプトの概要：
    アートボードごとにページ番号を自動挿入します。
    "_pagenumber" レイヤーに既存のテキストがある場合、そのフォント・フォントサイズ・行揃えを流用します。
    揃えは右固定です。

    処理の流れ：
    1. UIダイアログで余白・開始番号を設定
    2. "_pagenumber" レイヤーを取得（なければ作成）
    3. 既存のテキスト情報（フォントなど）を取得
    4. 各アートボードに右揃えでページ番号を挿入

    対象：開いているドキュメント内のすべてのアートボード
    更新日：2025-06-27
*/

function main() {
    if (app.documents.length > 0) {
        var dialog = new Window("dialog", "ページ番号を挿入");

        var marginPanel = dialog.add("panel", undefined, "余白");
        marginPanel.orientation = "row";
        marginPanel.add("statictext", undefined, "アートボード端からの距離：");
        var marginInput = marginPanel.add("edittext", undefined, "0.25");
        marginInput.characters = 5;
        marginPanel.add("statictext", undefined, "インチ");

        var startNumberPanel = dialog.add("panel", undefined, "開始番号");
        startNumberPanel.orientation = "row";
        startNumberPanel.add("statictext", undefined, "開始番号：");
        var startNumberInput = startNumberPanel.add("edittext", undefined, "1");
        startNumberInput.characters = 5;

        var okButton = dialog.add("button", undefined, "実行");
        dialog.alignChildren = "fill";

        okButton.onClick = function() {
            insertPageNumbers();
            dialog.close();
        };

        dialog.center();
        dialog.show();

        function insertPageNumbers() {
            var doc = app.activeDocument;

            var pageNumberLayer;
            try {
                pageNumberLayer = doc.layers["_pagenumber"];
            } catch (e) {
                pageNumberLayer = doc.layers.add();
                pageNumberLayer.name = "_pagenumber";
            }

            // 選択されたテキストを保持（除外対象にする）
            var selection = doc.selection;
            var selectedText = null;
            if (selection.length == 1 && selection[0].typename === "TextFrame") {
                selectedText = selection[0];
            }

            // _pagenumberレイヤー内の他のテキストを削除
            for (var t = pageNumberLayer.textFrames.length - 1; t >= 0; t--) {
                var tf = pageNumberLayer.textFrames[t];
                if (tf !== selectedText) {
                    tf.remove();
                }
            }

            // フォントなどの初期値
            var baseFont = app.textFonts.getByName("MyriadPro-Regular");
            var baseSize = 12;
            var baseJustify = Justification.RIGHT;

            // 既存テキストからスタイルを流用
            var texts = pageNumberLayer.textFrames;
            if (texts.length > 0) {
                var refText = texts[0];
                if (refText.contents !== "") {
                    var attr = refText.textRange.characterAttributes;
                    if (attr.textFont != null) baseFont = attr.textFont;
                    if (attr.size > 0) baseSize = attr.size;
                    baseJustify = refText.textRange.paragraphAttributes.justification;
                }
            }

            var marginPoints = Number(marginInput.text) * 72;
            var startNumber = parseInt(startNumberInput.text, 10) || 1;

            for (var i = 0; i < doc.artboards.length; i++) {
                var ab = doc.artboards[i];
                var bounds = ab.artboardRect;
                var pageNumber = startNumber + i;

                var tf = pageNumberLayer.textFrames.add();
                tf.contents = String(pageNumber);

                // スタイル適用
                var txtAttr = tf.textRange.characterAttributes;
                txtAttr.textFont = baseFont;
                txtAttr.size = baseSize;

                var paraAttr = tf.textRange.paragraphAttributes;
                paraAttr.justification = Justification.RIGHT;

                // 右下に配置
                var right = bounds[2] - marginPoints;
                var bottom = bounds[3] + marginPoints + baseSize;
                tf.left = right;
                tf.top = bottom;
            }

            app.redraw();
        }
    } else {
        alert("ドキュメントが開かれていません");
    }
}

main();