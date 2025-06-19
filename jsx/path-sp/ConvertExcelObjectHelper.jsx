#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：ConvertExcelObjectHelper.jsx

  【概要】
  このスクリプトは、同一フォルダ内にある複数の処理スクリプト（サブスクリプト）を
  順番に実行するための統合実行スクリプトです。

  【処理の流れ】
  1. 現在のスクリプトが存在するフォルダを取得
  2. 指定されたサブスクリプトファイル（例：ColorToK100Converter.jsx など）を順次読み込み・実行
  3. エラーがあればアラートで通知し、成功時にはログ出力

  【実行対象スクリプト】
  - ColorToK100Converter.jsx
  - CenterLineFromRect.jsx
  - ConnectToRectangle.jsx
  - ConnectToAngularU.jsx
  - ConnectToLShape.jsx

  【更新履歴】
  - 2025-06-15 初版作成
  - 2025-06-15 コメント最適化・構造整理
*/

// 対象スクリプトを実行
function runScript(fileName) {
    var currentScript = File($.fileName);
    var currentFolder = currentScript.parent;
    var targetScript = File(currentFolder + "/" + fileName);

    if (targetScript.exists) {
        try {
            $.evalFile(targetScript);
            $.writeln("実行成功: " + fileName);
        } catch (e) {
            alert("実行中にエラーが発生: " + fileName + "\n" + e);
        }
    } else {
        alert("スクリプトが見つかりません: " + fileName);
    }
}

runScript("ColorToK100Converter.jsx");
runScript("CenterLineFromRectangle.jsx");
runScript("ConnectLine1_Rectangle.jsx");
runScript("ConnectLine2_AngularU.jsx");
runScript("ConnectLine3_LShape.jsx");