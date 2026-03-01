#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
### スクリプト名：

TextOutlineMemo.jsx

This script supports only Japanese language.

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択したテキストフレームの情報を取得してアウトライン化
- 情報を note プロパティに保存し、後から参照可能にする

### 主な機能：

- テキスト内容・フォント名・サイズ・行送り・カーニングなどの取得
- テキストをアウトライン化し、情報を note に書き込む

### 処理の流れ：

1. テキストフレームを選択
2. 各種情報を取得
3. テキストをアウトライン化
4. 取得した情報を outlined オブジェクトの note に格納

### note

https://note.com/dtp_tranist/n/n3e0f241508db

### 更新履歴：

- v1.0 (20240723) : 初期バージョン
- v1.1 (20250721) : ローカライズ

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

function main() {
    // ドキュメントが開かれていることを確認
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var selection = doc.selection;

        // 選択されたオブジェクトが存在するか確認
        if (selection.length > 0) {
            // 元の選択を保存
            var originalSelection = selection.slice();

            for (var i = 0; i < originalSelection.length; i++) {
                if (originalSelection[i].typename == "TextFrame") {
                    processTextFrame(originalSelection[i]);
                }
            }
        } else {
            alert("テキストオブジェクトが選択されていません。");
        }
    } else {
        alert("ドキュメントが開かれていません。");
    }
}

function processTextFrame(textFrame) {
    var textRange = textFrame.textRange;

    // テキスト情報を取得
    var content = textRange.contents;
    var fontName = textRange.characterAttributes.textFont.name;
    var fontSize = roundToTwoDecimals(textRange.characterAttributes.size);
    var leadingValue = roundToTwoDecimals(textRange.characterAttributes.leading);
    var kerningMethod = textRange.characterAttributes.kerningMethod;
    var proportionalMetrics = textRange.characterAttributes.proportionalMetrics;
    var trackingValue = textRange.characterAttributes.tracking;
    var orientation = textFrame.orientation == TextOrientation.VERTICAL ? "縦組み" : "横組み";

    // カーニングメソッドの文字列表現を取得
    var kerningMethodText;
    switch (kerningMethod) {
        case AutoKernType.AUTO:
            kerningMethodText = "メトリクス";
            break;
        case AutoKernType.METRICSROMANONLY:
            kerningMethodText = "和文等幅";
            break;
        case AutoKernType.OPTICAL:
            kerningMethodText = "オプティカル";
            break;
        default:
            kerningMethodText = "なし";
    }

    // プロポーショナルメトリクスの文字列表現を取得
    var proportionalMetricsText = proportionalMetrics ? "true" : "false";

    // 座標を取得
    var position = textFrame.position;
    var x = roundToTwoDecimals(position[0]);
    var y = roundToTwoDecimals(position[1]);

    // メモ用のテキストを作成
    var memoText = "文字列：\n" + content + "\n\n" +
                   "フォント：\n" + fontName + "\n\n" +
                   "フォントサイズ：\n" + fontSize + "\n\n" +
                   "行送り：\n" + leadingValue + "\n\n" +
                   "カーニング：\n" + kerningMethodText + "\n\n" +
                   "プロポーショナルメトリクス：\n" + proportionalMetricsText + "\n\n" +
                   "トラッキング：\n" + trackingValue + "\n\n" +
                   "組み方向：\n" + orientation + "\n\n" +
                   "座標：\nX = " + x + ", Y = " + y;

    // 元の選択をクリアし、現在のテキストフレームを選択
    app.activeDocument.selection = null;
    textFrame.selected = true;

    // テキストをアウトライン化
    textFrame.createOutline();

    // アウトライン化したオブジェクトを取得
    var outlinedObject = app.activeDocument.selection[0];

    // メモに情報を入力
    outlinedObject.note = memoText;
}

function roundToTwoDecimals(value) {
    return Math.round(value * 100) / 100;
}

// スクリプトを実行

main();