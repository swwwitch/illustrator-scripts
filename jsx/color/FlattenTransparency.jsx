#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイナミックアクションを使って、ドキュメントの透明部分を分割・統合します。

詳細は README を参照してください。

### Overview

Flattens the transparency in the document by running a dynamic action.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenTransparency";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenTransparency.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenTransparency.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function act_FlattenTransparency () {

      var SET_NAME = "FlattenTransparency";
      var ACTION_NAME = "Flatten";

      var str = '/version 3 ' +
      '/name [ 19 ' +
      '466c617474656e5472616e73706172656e6379 ' +
      '] ' +
      '/isOpen 1 ' +
      '/actionCount 1 ' +
      '/action-1 { ' +
      '/name [ 7 ' +
      '466c617474656e ' +
      '] ' +
      '/keyIndex 0 ' +
      '/colorIndex 0 ' +
      '/isOpen 1 ' +
      '/eventCount 1 ' +
      '/event-1 { ' +
      '/useRulersIn1stQuadrant 0 ' +
      '/internalName (ai_plugin_flatten_transparency) ' +
      '/localizedName [ 30 ' +
      'e9808fe6988ee983a8e58886e38292e58886e589b2e383bbe7b5b1e59088 ' +
      '] ' +
      '/isOpen 1 ' +
      '/isOn 1 ' +
      '/hasDialog 1 ' +
      '/showDialog 0 ' +
      '/parameterCount 5 ' +
      '/parameter-1 { ' +
      '/key 1920169082 ' +
      '/showInPalette 4294967295 ' +
      '/type (integer) ' +
      '/value 100 ' +
      '} ' +
      '/parameter-2 { ' +
      '/key 1919253100 ' +
      '/showInPalette 4294967295 ' +
      '/type (unit real) ' +
      '/value 1200.0 ' +
      '/unit 592342629 ' +
      '} ' +
      '/parameter-3 { ' +
      '/key 1869902968 ' +
      '/showInPalette 4294967295 ' +
      '/type (boolean) ' +
      '/value 1 ' +
      '} ' +
      '/parameter-4 { ' +
      '/key 1869902699 ' +
      '/showInPalette 4294967295 ' +
      '/type (boolean) ' +
      '/value 1 ' +
      '} ' +
      '/parameter-5 { ' +
      '/key 1667463282 ' +
      '/showInPalette 4294967295 ' +
      '/type (boolean) ' +
      '/value 1 ' +
      '} ' +
      '} ' +
      '}';

      var f = new File('~/ScriptAction.aia');
      if (!f.open('w')) throw new Error('Failed to open action file for writing: ' + f.fsName);
      f.write(str);
      f.close();

      app.loadAction(f);
      try { f.remove(); } catch (_) {}

      app.doScript(ACTION_NAME, SET_NAME, false);
      app.unloadAction(SET_NAME, '');
    }

    act_FlattenTransparency();

})();
