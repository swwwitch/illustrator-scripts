#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

常駐パレットから BridgeTalk でメインエンジンへ送り、そこで一時アクション（.aia）を読み込み・実行・解除する再利用テンプレートです。
TemporaryAction.jsx と同じ動作を、toString() で送れる形（JSDoc・コメントなし、ASCII のみ、自己完結）にしています。

### Overview

A reusable template for palettes that send work to the main engine via BridgeTalk, where a temporary action (.aia) is loaded, played and unloaded.
It behaves like TemporaryAction.jsx, but is written so it survives toString() (no JSDoc or comments, ASCII only, self-contained).

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TemporaryActionWorker";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、パレットのワーカー関数の並びの「前」（WORKER_FUNCS を定義する位置より前）に貼る。
    //    ワーカー識別子は workerToActionHex / workerBuildActionNameLines / workerUnloadTemporaryActionSet /
    //    workerLoadTemporaryActionSet / workerRunTemporaryAction。一覧は TEMPORARY_ACTION_WORKER_FUNCS
    //    パレット側だけで使うのは buildTemporaryActionWorkerArg / buildTemporaryActionWorkerCall
    // 2. 送る関数の一覧に連結する（var の代入順に注意。このパーツより後ろで連結する）
    //      var WORKER_FUNCS = [workerA, workerB].concat(TEMPORARY_ACTION_WORKER_FUNCS);
    //      （push 方式なら WORKER_FUNCS.push.apply(WORKER_FUNCS, TEMPORARY_ACTION_WORKER_FUNCS);）
    // 3. ワーカー内で使う（アクション定義もワーカー内で組むとき）:
    //      function workerDoSomething() {
    //          var src = ["/version 3"].concat(workerBuildActionNameLines("", "MySet"), [ … ]).join(String.fromCharCode(10));
    //          return workerRunTemporaryAction(src, "MySet", "myAction") ? "OK" : "ERR:action";
    //      }
    //    パレット側で組んで送るとき（定義に \n \t や日本語があっても壊れない）:
    //      delegate(buildTemporaryActionWorkerCall(actionSource, "MySet", "myAction") + ' ? "OK" : "ERR:action"');
    //    何度も実行するとき: workerLoadTemporaryActionSet → try { app.doScript(…) } finally { workerUnloadTemporaryActionSet }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。戻り値はマーカー文字列にして返すこと
    //    （BridgeTalk の戻り値は文字列。true/false は "true"/"false" で届く）
    // 5. ワーカー関数を書き換えるときの決まり（toString() で送るため）:
    //    JSDoc・コメントを本体にも直前直後にも置かない（直後のコメントは取り込まれ、改行が落ちると後ろを潰す）。
    //    各文はセミコロンで終える。ASCII のみ（日本語は化ける）。文字列に \ を書かない（タブは String.fromCharCode(9)）。
    //    ほかの関数を呼ぶのは同じ一覧のワーカー関数だけ。関数の終わりの } は単独の行に置く。
    //    パーツ前後の var 文は、コメントがワーカーに取り込まれないための区切りを兼ねる

    // 一時アクション・ワーカー版（再利用パーツ） / Temporary action, worker version (reusable)

    /**
     * メインエンジンへ送る文字列を、呼び出し式に埋め込める引数リテラルにする（パレット側で使う）
     * @param {string} sourceText - 送る文字列
     * @returns {string} decodeURIComponent("…") の形の式
     */
    function buildTemporaryActionWorkerArg(sourceText) {
        return 'decodeURIComponent("' + encodeURIComponent(String(sourceText)) + '")';
    }

    /**
     * workerRunTemporaryAction の呼び出し式を組み立てる（パレット側で使う。評価結果は true / false）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {string} メインエンジンで評価する呼び出し式
     */
    function buildTemporaryActionWorkerCall(actionSource, setName, actionName) {
        return "workerRunTemporaryAction("
            + buildTemporaryActionWorkerArg(actionSource) + ", "
            + buildTemporaryActionWorkerArg(setName) + ", "
            + buildTemporaryActionWorkerArg(actionName) + ")";
    }

    /* 送るワーカー関数の一覧（関数宣言は巻き上がるのでここで参照できる） / Worker functions to send (declarations are hoisted) */
    var TEMPORARY_ACTION_WORKER_FUNCS = [
        workerToActionHex,
        workerBuildActionNameLines,
        workerUnloadTemporaryActionSet,
        workerLoadTemporaryActionSet,
        workerRunTemporaryAction
    ];

    function workerToActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    function workerBuildActionNameLines(indent, nameText, fieldName) {
        var nameHex = workerToActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + String.fromCharCode(9) + nameHex,
            indent + "]"
        ];
    }

    function workerUnloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
        }
    }

    function workerLoadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            workerUnloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("workerLoadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { }
            try { actionFile.remove(); } catch (removeError) { }
        }
    }

    function workerRunTemporaryAction(actionSource, setName, actionName) {
        if (!workerLoadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("workerRunTemporaryAction: " + e);
            return false;
        } finally {
            workerUnloadTemporaryActionSet(setName);
        }
    }

    /* 区切りの文（直後のコメントが最後のワーカーに取り込まれるのを防ぐ） / Separator statement: keeps the comments below out of the last worker */
    var TEMPORARY_ACTION_WORKER_END = true;

    // 一時アクション・ワーカー版（再利用パーツ）ここまで / End of the reusable temporary action worker

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Select objects first." },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    /**
     * 現在の UI 言語のラベルを返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} ラベル
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // デモ / Demo
    // =========================================
    var DEMO_SET_NAME = "TemporaryActionWorkerDemo";  /* デモのアクションセット名 / demo action set name */
    var DEMO_ACTION_NAME = "alignHCenter";            /* デモのアクション名 / demo action name */

    /**
     * 整列パネルの「水平方向中央に整列」を1件だけ持つアクション定義を返す（パレット側で組む例）
     * @returns {string} アクション定義のテキスト
     */
    function buildDemoAlignAction() {
        return ["/version 3"].concat(
            workerBuildActionNameLines("", DEMO_SET_NAME),
            ["/isOpen 1", "/actionCount 1", "/action-1 {"],
            workerBuildActionNameLines("\t", DEMO_ACTION_NAME),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (ai_plugin_alignPalette)"
            ],
            workerBuildActionNameLines("\t\t", "整列", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key 1954115685",
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (enumerated)"
            ],
            workerBuildActionNameLines("\t\t\t", "水平方向中央に整列"),
            [
                "\t\t\t/value 2",
                "\t\t}",
                "\t}",
                "}",
                ""
            ]
        ).join("\n");
    }

    /**
     * 送る関数を toString() で連結し、BridgeTalk の本文と同じ形（decodeURIComponent で包んだ eval）にする
     * @param {Function[]} workerFuncs - 送るワーカー関数
     * @param {string} callExpression - 末尾で評価する呼び出し式
     * @returns {string} メインエンジンで評価する本文
     */
    function buildDemoBridgeBody(workerFuncs, callExpression) {
        var sourceParts = [];
        for (var i = 0; i < workerFuncs.length; i++) {
            sourceParts.push(String(workerFuncs[i]).replace(/\r\n?/g, "\n"));
        }
        var evalSource = sourceParts.join("\n") + "\n" + callExpression + ";";
        return 'eval(decodeURIComponent("' + encodeURIComponent(evalSource) + '"));';
    }

    if (!app.documents.length || !app.activeDocument.selection.length) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }
    var demoBody = buildDemoBridgeBody(
        TEMPORARY_ACTION_WORKER_FUNCS,
        buildTemporaryActionWorkerCall(buildDemoAlignAction(), DEMO_SET_NAME, DEMO_ACTION_NAME) + ' ? "OK" : "ERR"'
    );
    /* 常駐パレットなら bridge.body = demoBody で送る。デモはこのエンジンのグローバルスコープで評価し、
       ワーカーが自己完結していること（このクロージャの関数に頼らないこと）を確かめる
       In a palette, send it as bridge.body; the demo evaluates it in global scope here to prove the workers are self-contained */
    var demoResult = String(new Function("return " + demoBody.replace(/;$/, ""))());
    if (demoResult !== "OK") {
        alert(getLabel(LABELS.alert.actionFailed));
    }

})();
