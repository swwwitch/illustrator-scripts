#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.DistributeRL
Meta-Default-language=ja
Name=右端を固定して左方向に等間隔で再配置
Name[en]=Keep the right edge fixed and redistribute leftward
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.2
Tags=alignment,整列
X-com.swwwitch-Script-Name=DistributeRL
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeRL.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeRL.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/single-function/DistributeRL.jsx
Description-BEGIN
[ja]
横並びに選択した複数オブジェクトのうち最も右のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ左方向へ等間隔に再配置します。
[en]
Keeps the rightmost object of a horizontal selection fixed and redistributes the rest to the left at intervals of the Keyboard Increment.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

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
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /**
     * 左端X（position[0]）で並べ替えた新しい配列を返す
     * @param {PageItem[]} targetObjects - 並べ替える対象のオブジェクト
     * @param {boolean} rightFirst - true なら右から左（降順）、false なら左から右（昇順）
     * @returns {PageItem[]} 並べ替えた新しい配列
     */
    function sortByLeftEdge(targetObjects, rightFirst) {
        var sortedObjects = [];
        for (var i = 0; i < targetObjects.length; i++) sortedObjects.push(targetObjects[i]);
        sortedObjects.sort(function (itemA, itemB) {
            var leftDelta = itemA.position[0] - itemB.position[0];
            return rightFirst ? -leftDelta : leftDelta;
        });
        return sortedObjects;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 最も右のオブジェクトを固定し、以降をキー入力の値ずつ左へ等間隔に再配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length < 1) return;

        var selectedObjects = app.activeDocument.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length < 2) return;

        /* 環境設定［一般］の「キー入力」（cursorKeyLength、pt）を移動幅に使う / Keyboard Increment in points */
        var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength");

        /* 最も右を固定し、以降を keyboardIncrementPt ずつ左へ等間隔配置 / Keep the rightmost one, move the rest left step by step */
        var objectsRightToLeft = sortByLeftEdge(selectedObjects, true);
        for (var i = 1; i < objectsRightToLeft.length; i++) {
            objectsRightToLeft[i].translate(-i * keyboardIncrementPt, 0);
        }
    }

    main();

})();
