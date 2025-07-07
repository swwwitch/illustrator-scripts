#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ConvertExcelObjectHelper.jsx

### 概要

- 同一フォルダ内にある複数のサブスクリプトを順番に実行する統合スクリプトです。
- Illustratorでのパーツ配置や補助処理を一括で実行する際に便利です。

### 主な機能

- フォルダ内の指定サブスクリプトを順番に呼び出し実行
- エラー発生時にはアラートで通知
- 成功時にはコンソールにログ出力
- Illustrator ExtendScript (ES3) 準拠

### 処理の流れ

1. 現在のスクリプトが存在するフォルダを取得
2. 各サブスクリプトファイルを順次読み込み・実行
3. エラーがあればアラート表示、成功時はログ出力

### 更新履歴

- v1.0.0 (20250615) : 初期バージョン
- v1.0.1 (20250615) : コメント最適化・構造整理

---

### Script Name:

ConvertExcelObjectHelper.jsx

### Overview

- An integrated script to execute multiple sub-scripts in the same folder sequentially.
- Useful for batch operations like part arrangement or support processes in Illustrator.

### Main Features

- Sequentially execute specified sub-scripts in the same folder
- Alert notification on error
- Console log output on success
- Compatible with Illustrator ExtendScript (ES3)

### Process Flow

1. Get the folder where the current script exists
2. Sequentially load and execute each sub-script file
3. Show alert if any error occurs, log success to console

### Update History

- v1.0.0 (20250615): Initial version
- v1.0.1 (20250615): Optimized comments and refined structure
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