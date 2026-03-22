#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * スクリプト名 / Script Name: PresetManagerNoDialogFull.jsx
 * 説明 / Description:
 *   Illustrator の各種環境設定を一括で適用します。ダイアログは表示しません。
 *   Apply a curated set of Illustrator preferences at once. No dialogs are shown.
 *
 * 注意 / Notes:
 *   - バージョン差異により一部キーは無視される場合があります。
 *   - Some preference keys may not exist depending on your Illustrator version; such keys are safely ignored.
 *
 * 更新履歴 / Update History:
 *   - v1.0 (2025-08-18): 初期バージョン / Initial release
 */

with(app.preferences) {

    /*
     * 一般 / General
     * General app-level preferences.
     */

    // キー入力
    setRealPreference("cursorKeyLength", 0.2835); // 0.1mm

    // 詳細なツールヒントを表示
    setBooleanPreference("showRichToolTips", false); // true=表示, false=非表示

    // 裁ち落とし部分に「裁ち落としを印刷」生成AIボタンを表示
    setBooleanPreference("enablePrintBleedWidget", false); // true=表示, false=非表示

    /*
     * 選択・表示 / Selection and Display
     * Selection, zoom, and display-related preferences.
     */

    // ロックまたは非表示オブジェクトをアートボードと一緒に移動
    setBooleanPreference("moveLockedAndHiddenArt", true); // true=移動可能, false=移動不可

    /*
     * テキスト / Text
     * Text behavior and default content preferences.
     */

    // 新規エリア内文字の自動サイズ調整
    setBooleanPreference("text/autoSizing", true); // true=有効, false=無効

    // テキスト > 新規テキストオブジェクトにサンプルテキストを割り付け OFF
    setBooleanPreference("text/fillWithDefaultText", false);
    setBooleanPreference("text/fillWithDefaultTextJP", false);

    // 見つからない字形の保護を有効にする
    setBooleanPreference("text/doFontLocking", false); // true=有効, false=無効

    // 選択された文字の異体字を表示
    setBooleanPreference("text/enableAlternateGlyph", false); // true=有効, false=無効

    // テキスト > 入力中に箇条書きリストのレベルを自動更新 OFF
    setBooleanPreference("text/enableListAutoDetection", false);

    /*
     * スマートガイド / Smart Guides
     * Smart Guides visibility and behavior.
     */

    // オブジェクトのハイライト表示
    setBooleanPreference("smartGuides/showObjectHighlighting", false); // true=有効, false=無効

    /*
     * UI / User Interface
     * Canvas color and UI appearance.
     */

    // カンバスカラー
    setIntegerPreference("uiCanvasIsWhite", 1); // 0=UIに合わせる,1=ホワイト

    /*
     * パフォーマンス / Performance
     * Performance-related toggles such as undo depth and live editing.
     */

    // ヒストリー数
    setIntegerPreference("maximumUndoDepth", 50); // default:100

    // リアルタイムの描画と編集
    setBooleanPreference("LiveEdit_State_Machine", true); // true=有効, false=無効

    /*
     * ファイル管理 / File Management
     * Font activation and file handling preferences.
     */

    // Adobe Fontsを自動アクティベート
    setBooleanPreference("AutoActivateMissingFont", true); // true=有効, false=無効

    /*
     * ブラックのアピアランス / Black Appearance
     * Preserve accurate black on screen and export.
     */

    // ブラックのアピアランス > すべてのブラックを正確に表示／出力
    setIntegerPreference("blackPreservation/Onscreen", 0); //スクリーン
    setIntegerPreference("blackPreservation/Export", 0); //プリント／書き出し

    /*
     * 環境設定以外 / Other Preferences
     * Other toggles not strictly under the main Preferences panels.
     */

    // 初期設定のファイルの場所 コンピューター 1、Creative Cloud true
    setBooleanPreference("AdobeSaveAsCloudDocumentPreference", false);

    // コンテキストタスクバー
    setBooleanPreference('ContextualTaskBarEnabled', false); // true=有効, false=無効

    // パスファインダー適用時のアラートをOFFに
    setBooleanPreference('plugin/DontShowWarningAgain/ShowPathfinderGroupWarning', false); // true=有効, false=無効;

    // 任意設定群を一括適用するスイッチ（ユーザーの好みが分かれる設定）/ Toggle to apply optional preference set (user-dependent).
    var applyOptionalPrefs = true; // 任意設定を適用するか？
    if (applyOptionalPrefs) {

        // 最近使用したフォントの表示数
        setIntegerPreference("text/recentFontMenu/showNEntries", 0); // 0=非表示, 1以上=表示数

        // 選択範囲へズーム
        setBooleanPreference("zoomToSelection", false); // true=有効, false=無効

        // ドキュメントを開いていないときにホーム画面を表示
        setBooleanPreference("Hello/ShowHomeScreenWS", false); // true=表示, false=非表示

        // 以前の「新規ドキュメント」インターフェイスを使用
        setBooleanPreference("Hello/NewDoc", true); // true=有効, false=無効

        // ファイル管理 > オリジナルの編集にシステムデフォルトを使用 ON
        setBooleanPreference("useSysDefEdit", true);

        // アニメーションズーム
        setBooleanPreference("Performance/AnimZoom", false); // true=有効, false=無効
    }

    // 旧来の選択ヒット判定に寄せるオプション / Optional legacy-style hit-test preferences.
    var applyOptionalPrefsForOldSchool = false; // 任意設定を適用するか？
    if (applyOptionalPrefsForOldSchool) {

        // 現在の「オブジェクトの選択範囲をパスに制限」
        setIntegerPreference("hitShapeOnPreview", 0); // 0がON（true）、1がOFF（false）
        // 現在の「テキストをパスに制限」
        setIntegerPreference("hitTypeShapeOnPreview", 0); //

    } else {

        // 現在の「オブジェクトの選択範囲をパスに制限」
        setIntegerPreference("hitShapeOnPreview", 1); // 0がON（true）、1がOFF（false）
        // 現在の「テキストをパスに制限」
        setIntegerPreference("hitTypeShapeOnPreview", 1); //


    }

    var needsRestart = true;
    if (needsRestart) {
        // alert('一部の設定を反映するには Illustrator の再起動が必要です。');

        // 定規
        setIntegerPreference("rulerType", 1); // mm

        // ユーザーインターフェイスの明るさ
        setRealPreference("uiBrightness", 1); // 0=暗, 0.5=やや暗め, 0.50999999046326=やや明るめ, 1=明るい

        // ［シェイプ形成ツール］の［次のカラー］を「オブジェクト」に
        setIntegerPreference("Planar/MergeTool/PaintFills", 0); // 0：オブジェクト、1：スウォッチ

    }

}