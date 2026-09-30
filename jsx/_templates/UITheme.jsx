#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

UI がダークテーマかどうかを判定する再利用テンプレートです。自作描画する部品（StepperButtons・LinkToggle など）の配色に使います。
Illustrator・InDesign のどちらでも使えます。

### Overview

A reusable template that tells whether the UI uses a dark theme, for the colors of custom-drawn parts such as StepperButtons and LinkToggle.
Works in both Illustrator and InDesign.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UITheme";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // =========================================
    // デモ / Demo
    // =========================================
    alert(isDarkUI() ? "Dark UI" : "Light UI");

})();
