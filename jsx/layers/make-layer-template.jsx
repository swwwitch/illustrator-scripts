/*
  
### スクリプト名：

make-layer-template.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択しているレイヤーを「テンプレート」属性（ロック、印刷不可）に変更
- ダイナミックアクションを使い、ダイアログなしで実行

### 更新履歴：

- v1.0 (20240721) : 初期バージョン

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.0";

function setLayTmplAttr() {
  /* 
    アクション文字列の定義
    Define the action string for setting template layer properties
  */
  var actionString = 
    '/version 3' +'/name [ 5' +' 6c61796572' +' ]' +'/isOpen 1' +'/actionCount 1' +'/action-1 {' +' /name [ 24' +' 6368616e67652d746f2d74656d706c6174652d6c61796572' +' ]' +' /keyIndex 0' +' /colorIndex 0' +' /isOpen 1' +' /eventCount 1' +' /event-1 {' +' /useRulersIn1stQuadrant 0' +' /internalName (ai_plugin_Layer)' +' /localizedName [ 9' +' e8a1a8e7a4ba203a20' +' ]' +' /isOpen 1' +' /isOn 1' +' /hasDialog 1' +' /showDialog 0' +' /parameterCount 10' +
    ' /parameter-1 {' +' /key 1836411236' +' /showInPalette 4294967295' +' /type (integer)' +' /value 4' +' }' +
    ' /parameter-2 {' +' /key 1851878757' +' /showInPalette 4294967295' +' /type (ustring)' +' /value [ 36' +' e383ace382a4e383a4e383bce38391e3838de383aae382aae38397e382b7e383' +' a7e383b3' +' ]' +' }' +
    ' /parameter-3 {' +' /key 1953068140' +' /showInPalette 4294967295' +' /type (ustring)' +' /value [ 13' +' 6f75746c696e65642d74657874' +' ]' +' }' +
    ' /parameter-4 {' +' /key 1953329260' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 1' +' }' +
    ' /parameter-5 {' +' /key 1936224119' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 1' +' }' +
    ' /parameter-6 {' +' /key 1819239275' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 1' +' }' +
    ' /parameter-7 {' +' /key 1886549623' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 1' +' }' +
    ' /parameter-8 {' +' /key 1886547572' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 0' +' }' +
    ' /parameter-9 {' +' /key 1684630830' +' /showInPalette 4294967295' +' /type (boolean)' +' /value 1' +' }' +
    ' /parameter-10 {' +' /key 1885564532' +' /showInPalette 4294967295' +' /type (unit real)' +' /value 50.0' +' /unit 592474723' +' }' +' }' +
    '}';

  /* 
    アクションファイルを作成して書き込み
    Create and write temporary action file
  */
  var actionFile = new File('~/ScriptAction.aia');
  actionFile.open('w');
  actionFile.write(actionString);
  actionFile.close();

  /* 
    アクションを読み込んで実行し、削除
    Load, execute, and remove the action
  */
  app.loadAction(actionFile);
  actionFile.remove();
  app.doScript("change-to-template-layer", "layer", false); // action name, set name
  app.unloadAction("layer", ""); // set name
}

/* 
  メイン処理の実行
  Run main process
*/
setLayTmplAttr();