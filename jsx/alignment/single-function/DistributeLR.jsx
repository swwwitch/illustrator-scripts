#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.DistributeLR
Meta-Default-language=ja
Name=左端を固定して右方向に等間隔で再配置
Name[en]=Keep the left edge fixed and redistribute rightward
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.2
Tags=alignment,整列
X-com.swwwitch-Script-Name=DistributeLR
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeLR.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeLR.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/single-function/DistributeLR.jsx
Description-BEGIN
[ja]
横並びに選択した複数オブジェクトのうち最も左のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ右方向へ等間隔に再配置します。
[en]
Keeps the leftmost object of a horizontal selection fixed and redistributes the rest to the right at intervals of the Keyboard Increment.
Description-END
SCRIPTMETA-END
*/

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
     * 最も左のオブジェクトを固定し、以降をキー入力の値ずつ右へ等間隔に再配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length < 1) return;

        var selectedObjects = app.activeDocument.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length < 2) return;

        /* 環境設定［一般］の「キー入力」（cursorKeyLength、pt）を移動幅に使う / Keyboard Increment in points */
        var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength");

        /* 最も左を固定し、以降を keyboardIncrementPt ずつ右へ等間隔配置 / Keep the leftmost one, move the rest right step by step */
        var objectsLeftToRight = sortByLeftEdge(selectedObjects, false);
        for (var i = 1; i < objectsLeftToRight.length; i++) {
            objectsLeftToRight[i].translate(i * keyboardIncrementPt, 0);
        }
    }

    main();

})();
