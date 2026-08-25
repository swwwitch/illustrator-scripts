#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトの見た目を、固定名のグラフィックスタイルとして登録します。
同名スタイルが既に存在する場合は削除してから登録するため、繰り返し実行しても重複しません。

詳細は README を参照してください。

### Overview

Registers the appearance of the selected object as a graphic style with a fixed name.
An existing style with the same name is removed first, so repeated runs never create duplicates.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "register-temp-style";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/register-temp-style.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/register-temp-style.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 登録するスタイル名 / Style name to register */
var TEMP_STYLE_NAME = "temp_style";

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
    if (selectedItems.length !== 1) {
        return;
    }

    // 既存の temp_style があれば削除
    try { graphicStyles.getByName(TEMP_STYLE_NAME).remove(); }
    catch (e) { }

    // アクションをロード → 実行 → アンロード
    var beforeCount = graphicStyles.length;
    loadForceNewGraphicStyleAction();
    runForceNewGraphicStyleAction();
    unloadForceNewGraphicStyleAction();

    // アクションでスタイルが追加された場合のみ、末尾を temp_style に改名
    if (graphicStyles.length > beforeCount) {
        graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
    }
})();
