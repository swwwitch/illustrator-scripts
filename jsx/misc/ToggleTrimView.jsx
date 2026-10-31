#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ToggleTrimView
Name=トリミング表示を切り替えてガイドを表示
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0
Description-BEGIN
トリミング表示を切り替え、ガイドを表示します。

Toggles Trim View and shows the guides.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

トリミング表示を切り替え、ガイドを表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ToggleTrimView.md

### Overview

Toggles Trim View and shows the guides.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ToggleTrimView.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ToggleTrimView";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "トリミング表示を切り替えてガイドを表示"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ToggleTrimView.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ToggleTrimView.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ToggleTrimView.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    app.executeMenuCommand('TrimView');
    app.executeMenuCommand('showguide');

})();
