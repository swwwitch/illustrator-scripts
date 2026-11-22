#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ScriptMeta
Meta-Default-language=ja
Name=SCRIPTMETA から基本情報を読む部品
Name[en]=Read the basic info from SCRIPTMETA
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-11-21
Tags=misc,その他
X-com.swwwitch-Script-Name=ScriptMeta
X-com.swwwitch-Updated=2026-11-22
Description-BEGIN
[ja]
スクリプト先頭の SCRIPTMETA を実行時に読み、getScriptMetaValue() で項目を直接返す再利用テンプレートです。設定の保存名に使う SCRIPT_NAME だけは変数で用意します。
[en]
A reusable template that reads the SCRIPTMETA block at the top of the script at run time and returns its fields directly through getScriptMetaValue(). Only SCRIPT_NAME, the key for saved settings, is kept as a variable.
Description-END
SCRIPTMETA-END
*/

// 【移植手順 / How to port】
// 1. スクリプトの先頭に SCRIPTMETA を書く（既存のスクリプトからコピーして Script-ID・Name・Version・
//    X-com.swwwitch-Script-Name などを書き換える）。Name・Name[en]・Tags・Source-URL は helpers/sync-scriptmeta.py が入れる
// 2. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、IIFE の外（スクリプトの最上位）の
//    「基本情報 / Basic info」の区切りの下に貼る。中身は helpers/sync-scriptmeta.py がこのファイルに合わせて書き直す
// 3. SCRIPT_NAME（設定の保存名・ダイアログの位置のキー）はこれまでどおり変数で使える。
//    ほかの項目は変数を作らず getScriptMetaValue("Version") のように直接読む（VersionFooter の部品もそうする）。
//    バージョンを表示するときは "v" + getScriptMetaValue("Version")
// 4. バージョン・更新日・説明を変えるときは SCRIPTMETA の Version・X-com.swwwitch-Updated・Description を書き換える

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    alert(SCRIPT_NAME + " v" + getScriptMetaValue("Version") + "\n" + getScriptMetaValue("Name") + "\n" + getScriptMetaValue("X-com.swwwitch-Updated"));
})();
