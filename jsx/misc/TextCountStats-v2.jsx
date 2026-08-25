/*

### 概要

トリミング表示を切り替え、ガイドを表示します。

詳細は README を参照してください。

### Overview

Toggles Trim View and shows the guides.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextCountStats-v2";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStats-v2.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStats-v2.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

app.executeMenuCommand('TrimView');
app.executeMenuCommand('showguide');
