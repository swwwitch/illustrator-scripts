#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* 環境設定を変更 */
function main() {
    // 詳細なツールヒントを表示
    // true=表示, false=非表示
    // app.preferences.setBooleanPreference("showRichToolTips", true); // default
    app.preferences.setBooleanPreference("showRichToolTips", false);

    // ドキュメントを開いていないときにホーム画面を表示
    // true=表示, false=非表示
    // app.preferences.setBooleanPreference("Hello/ShowHomeScreenWS", true); // default
    app.preferences.setBooleanPreference("Hello/ShowHomeScreenWS", false);

    // 以前の「新規ドキュメント」インターフェイスを使用
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("Hello/NewDoc", false); // default
    app.preferences.setBooleanPreference("Hello/NewDoc", true);

    // 裁ち落とし部分に「裁ち落としを印刷」生成AIボタンを表示
    // true=表示, false=非表示
    // app.preferences.setBooleanPreference("enablePrintBleedWidget", true); // default
    app.preferences.setBooleanPreference("enablePrintBleedWidget", false);

    /* ==============
    選択・表示 / Selection and Display
    */

    // ロックまたは非表示オブジェクトをアートボードと一緒に移動
    // true=移動可能, false=移動不可
    // app.preferences.setBooleanPreference("moveLockedAndHiddenArt", false); // default
    app.preferences.setBooleanPreference("moveLockedAndHiddenArt", true);

    // 選択範囲へズーム
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("zoomToSelection", true); // default
    app.preferences.setBooleanPreference("zoomToSelection", false);

    /* ==============
    テキスト / Text
    */

    // 新規エリア内文字の自動サイズ調整
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("text/autoSizing", false); // default
    app.preferences.setBooleanPreference("text/autoSizing", true);

    // 最近使用したフォントの表示数
    // 0=非表示, 1以上=表示数
    // app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 10); // default
    app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 0);

    // 見つからない字形の保護を有効にする
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("text/doFontLocking", true); // default
    app.preferences.setBooleanPreference("text/doFontLocking", false);

    // 選択された文字の異体字を表示
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("text/enableAlternateGlyph", true); // default
    app.preferences.setBooleanPreference("text/enableAlternateGlyph", false);

    /* ==============
    UI / User Interface
    */

    // カンバスカラー
    // 0=黒, 1=白
    // app.preferences.setIntegerPreference("uiCanvasIsWhite", 0); // default
    app.preferences.setIntegerPreference("uiCanvasIsWhite", 1);

    /* ==============
    パフォーマンス / Performance
    */

    // アニメーションズーム
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("Performance/AnimZoom", true); // default
    app.preferences.setBooleanPreference("Performance/AnimZoom", false);

    // ヒストリー数
    // app.preferences.setIntegerPreference("maximumUndoDepth", 100); // default
    app.preferences.setIntegerPreference("maximumUndoDepth", 50);

    // リアルタイムの描画と編集
    // true=有効, false=無効 
    // app.preferences.setBooleanPreference("LiveEdit_State_Machine", true); // default
    app.preferences.setBooleanPreference("LiveEdit_State_Machine", true);

    /* ==============
    ファイル管理 / File Management
    */

    // Adobe Fontsを自動アクティベート
    // true=有効, false=無効
    // app.preferences.setBooleanPreference("AutoActivateMissingFont", false); // default
    app.preferences.setBooleanPreference("AutoActivateMissingFont", true);

    // ファイルの保存先：クラウド/コンピューター
    // true=クラウド, false=コンピューター
    // app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", true); // default
    app.preferences.setBooleanPreference("AdobeSaveAsCloudDocumentPreference", false);
}

main();