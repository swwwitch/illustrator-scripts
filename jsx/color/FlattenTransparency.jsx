#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイナミックアクションを使って、ドキュメントの透明部分を分割・統合します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenTransparency.md

### Overview

Flattens the transparency in the document by running a dynamic action.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenTransparency.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenTransparency";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenTransparency.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenTransparency.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイナミックアクション / Dynamic action
    // =========================================

    /**
     * 「透明部分を分割・統合」をダイナミックアクションで実行する（ダイアログは出さない）
     * アクション定義を一時ファイルに書き出して読み込み、実行後に取り除く
     * @returns {void}
     */
    function runFlattenTransparencyAction() {
        var SET_NAME = "FlattenTransparency";
        var ACTION_NAME = "Flatten";

        var actionCode = [
            "/version 3",
            "/name [ 19",
            "466c617474656e5472616e73706172656e6379",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "/name [ 7",
            "466c617474656e",
            "]",
            "/keyIndex 0",
            "/colorIndex 0",
            "/isOpen 1",
            "/eventCount 1",
            "/event-1 {",
            "/useRulersIn1stQuadrant 0",
            "/internalName (ai_plugin_flatten_transparency)",
            "/localizedName [ 30",
            "e9808fe6988ee983a8e58886e38292e58886e589b2e383bbe7b5b1e59088",
            "]",
            "/isOpen 1",
            "/isOn 1",
            "/hasDialog 1",
            "/showDialog 0",
            "/parameterCount 5",
            "/parameter-1 {",
            "/key 1920169082",
            "/showInPalette 4294967295",
            "/type (integer)",
            "/value 100",
            "}",
            "/parameter-2 {",
            "/key 1919253100",
            "/showInPalette 4294967295",
            "/type (unit real)",
            "/value 1200.0",
            "/unit 592342629",
            "}",
            "/parameter-3 {",
            "/key 1869902968",
            "/showInPalette 4294967295",
            "/type (boolean)",
            "/value 1",
            "}",
            "/parameter-4 {",
            "/key 1869902699",
            "/showInPalette 4294967295",
            "/type (boolean)",
            "/value 1",
            "}",
            "/parameter-5 {",
            "/key 1667463282",
            "/showInPalette 4294967295",
            "/type (boolean)",
            "/value 1",
            "}",
            "}",
            "}"
        ].join(" ");

        /* 失敗は従来どおり例外にして Illustrator のエラー表示に任せる / Throw on failure so Illustrator reports it, as before */
        if (!runTemporaryAction(actionCode, SET_NAME, ACTION_NAME)) {
            throw new Error("Could not run the action \"" + ACTION_NAME + "\".");
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    runFlattenTransparencyAction();

})();
