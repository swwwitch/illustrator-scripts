#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要 / Overview

- オブジェクトとテキストを同時に選択して実行すると、オブジェクトの見た目をグラフィックスタイルに登録し、テキストの文字列をそのスタイル名にするスクリプト。
- 同名スタイルが既に存在する場合は削除してから上書き登録するので、繰り返し実行しても重複しない。
- Registers the appearance of the selected object as a graphic style, using the selected text's content as the style name.
- If a style with the same name already exists, it is removed first so repeated runs do not create duplicates.

### 処理の流れ / Process flow

1. 選択が「テキスト1点＋オブジェクト1点」でなければ何もせず終了 / Exit unless exactly one TextFrame and one non-text object are selected.
2. テキストから文字列を取得（前後の空白を除去、空なら終了） / Get the text content (trimmed); exit if empty.
3. 既存の同名スタイルを削除 / Remove the existing style with the same name, if any.
4. スタイル対象オブジェクトだけを選択し直す / Reselect only the style-target object (exclude the text's appearance).
5. 一時アクションをロード→実行→アンロードして、無名のグラフィックスタイルを末尾に追加 / Load, run, and unload a temporary action that appends an unnamed graphic style.
6. 末尾のスタイルを取得した文字列に改名 / Rename the last style to the text content.
7. 元の選択に戻す / Restore the original selection.

*/

// =========================================
// バージョンと設定 / Version & Settings
// =========================================

var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ヘルパー / Helpers
// =========================================

/* 前後の空白を除去（ES3 には String.trim が無い） / Trim leading & trailing whitespace */
function trimText(value) {
    return String(value).replace(/^\s+/, '').replace(/\s+$/, '');
}

// =========================================
// グラフィックスタイル関連 / Graphic Style Helpers
// =========================================

/* アクションを一時ファイルに書き出してロードする / Write & load the force-new-style action */
function loadForceNewGraphicStyleAction() {
    var actionData = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382 a4e383ab ] /value 1 } } }';

    var actionFile = new File(Folder.temp.fsName + '/__tmp_register_style.aia');
    actionFile.open('w');
    actionFile.write(actionData);
    actionFile.close();
    app.loadAction(actionFile);
    actionFile.remove();
}

/* ロード済みアクションを実行（選択オブジェクトをスタイル登録） / Run the loaded action on current selection */
function runForceNewGraphicStyleAction() {
    app.doScript('AddNewWithoutName', 'GraphicStyle', false);
}

/* アクションセットをアンロード / Unload the action set */
function unloadForceNewGraphicStyleAction() {
    app.unloadAction('GraphicStyle', '');
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var activeDoc = app.activeDocument;
    var graphicStyles = activeDoc.graphicStyles;

    var selectedItems = activeDoc.selection;
    if (selectedItems.length !== 2) {
        return;
    }

    // 選択を「テキスト」と「スタイル対象オブジェクト」に仕分け
    var textItem = null;
    var styleTarget = null;
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i].typename === 'TextFrame') {
            textItem = selectedItems[i];
        } else {
            styleTarget = selectedItems[i];
        }
    }
    if (!textItem || !styleTarget) {
        return;
    }

    // テキストからスタイル名を取得
    var styleName = trimText(textItem.contents);
    if (styleName === '') {
        return;
    }

    // 既存の同名スタイルがあれば削除
    try { graphicStyles.getByName(styleName).remove(); }
    catch (e) { }

    // スタイル対象オブジェクトだけを選択（テキストの見た目を含めない）
    activeDoc.selection = [styleTarget];

    // アクションをロード → 実行 → アンロード
    var beforeCount = graphicStyles.length;
    loadForceNewGraphicStyleAction();
    runForceNewGraphicStyleAction();
    unloadForceNewGraphicStyleAction();

    // アクションでスタイルが追加された場合のみ、末尾を取得した文字列に改名
    if (graphicStyles.length > beforeCount) {
        graphicStyles[graphicStyles.length - 1].name = styleName;
    }

    // 元の選択に戻す
    activeDoc.selection = selectedItems;
})();
