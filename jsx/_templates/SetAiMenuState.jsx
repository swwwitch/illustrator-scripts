#target illustrator

/*

### 概要

［表示］→［スマートガイド］や［ウィンドウ］→［コントロール］など、メニューに ✓ が付く項目の状態を読み、オン・オフどちらかに揃える再利用テンプレートです。
executeMenuCommand() は反転しかできないため、補助アプリ /Applications/SetAiMenuState.app がメニューの ✓ を読みます。

### Overview

A reusable template that reads checkmarked menu items such as View > Smart Guides or Window > Control and sets them on or off.
executeMenuCommand() can only toggle, so the helper app /Applications/SetAiMenuState.app reads the checkmarks.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetAiMenuState";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-09";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る。識別子は getAiMenuState / setAiMenuState
    // 2. 補助アプリは ai-scripts の helpers/SetAiMenuState.applescript から作る:
    //      osacompile -o /Applications/SetAiMenuState.app helpers/SetAiMenuState.applescript
    //    アクセシビリティの許可が要る。作り直すと許可もやり直し
    // 3. メニューの道筋は「>」でつないだ UI 上の項目名。英語版の Illustrator では英語名になる
    //    「ガイドを表示／隠す」のように名前が入れ替わる項目は「オフのときの名前|オンのときの名前」と書く:
    //      "表示>ガイド>ガイドを表示|ガイドを隠す"（「ガイドを隠す」が出ていれば on）
    // 4. 状態を読む: var states = getAiMenuState(["ウィンドウ>コントロール", "表示>グリッドにスナップ"]);
    //      states["ウィンドウ>コントロール"] が "on" / "off" / "notfound"。補助アプリが使えなければ null
    //    Illustrator はスクリプトの実行中にメニューを読ませないため、読む間だけ小さなダイアログボックスを出す
    // 5. 揃える: setAiMenuState([["表示>スマートガイド", "on"], ["表示>ポイントにスナップ", "off"]]);
    //      切り替わるのはスクリプトが終わった直後。切り替わった前提で処理を続けることはできない
    // 6. ポイントにスナップ・スマートガイド・グリフにスナップは補助アプリが要らない。下の「設定キーで読めるトグル」を使う:
    //      getPrefMenuToggle("glyphSnap") → true / false、setPrefMenuToggle("glyphSnap", true)
    //    環境設定キーで状態を読み、異なるときだけ executeMenuCommand() で反転する。スクリプトの中ですぐ切り替わる

    // メニューの ✓ を読む・揃える（再利用パーツ） / Read and set menu checkmarks (reusable)

    /**
     * 補助アプリ /Applications/SetAiMenuState.app に依頼を書いて起動する
     * @param {Array} requests - [メニューの道筋, "get"|"on"|"off"|"toggle"] の配列
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function launchAiMenuStateHelper(requests) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var helperAppPath = "/Applications/SetAiMenuState.app";
        var requestFilePath = "/tmp/set_ai_menu_state_request.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(helperAppPath).exists && !new File(helperAppPath).exists) return false;

        var lines = [];
        for (var i = 0; i < requests.length; i++) {
            lines.push(requests[i][0] + "\t" + requests[i][1]);
        }

        var requestFile = new File(requestFilePath);
        var written = false;
        try {
            requestFile.encoding = "UTF-8";
            requestFile.lineFeed = "Unix";
            if (requestFile.open("w")) {
                written = requestFile.write(lines.join("\n"));
            }
        } catch (e) {
        } finally {
            try { requestFile.close(); } catch (closeError) {}
        }
        return written && new File(helperAppPath).execute();
    }

    /**
     * メニュー項目の ✓ を読む。
     * スクリプトの実行中はメニューを読めないため、補助アプリが読み終えるまで小さなダイアログボックスを出して待つ
     * @param {string[]} menuPaths - 「表示>スマートガイド」の形のメニューの道筋
     * @returns {Object|null} 道筋をキーに "on" / "off" / "notfound" を持つオブジェクト。読めなかったら null
     */
    function getAiMenuState(menuPaths) {
        /* 補助アプリはこの名前の UI element を探してボタンを押す / The helper looks up a UI element with this name and presses its button */
        var waitDialogName = "SetAiMenuState";
        var resultFile = new File("/tmp/set_ai_menu_state_result.txt");
        var isJapanese = $.locale.indexOf("ja") === 0;

        /* 前回の結果を読まないように消しておく / Remove the previous result so it is never read by mistake */
        if (resultFile.exists) resultFile.remove();

        var requests = [];
        for (var i = 0; i < menuPaths.length; i++) requests.push([menuPaths[i], "get"]);
        if (!launchAiMenuStateHelper(requests)) return null;

        var waitDialog = new Window("dialog", waitDialogName);
        waitDialog.add("statictext", undefined, isJapanese ? "メニューの状態を確認中..." : "Checking menu states...");
        /* 補助アプリが動かないときは手で閉じる / Closed by hand when the helper cannot run */
        var cancelButton = waitDialog.add("button", undefined, isJapanese ? "キャンセル" : "Cancel");
        cancelButton.onClick = function () { waitDialog.close(); };
        waitDialog.show();

        if (!resultFile.exists) return null;
        var resultText = "";
        try {
            resultFile.encoding = "UTF-8";
            if (resultFile.open("r")) resultText = resultFile.read();
        } catch (e) {
        } finally {
            try { resultFile.close(); } catch (closeError) {}
        }

        /* 1行は「道筋<TAB>変更前<TAB>変更後」。error 行があれば失敗 / Each line is "path<TAB>before<TAB>after"; an error line means failure */
        var states = {};
        var resultLines = resultText.split("\n");
        for (var j = 0; j < resultLines.length; j++) {
            var fields = resultLines[j].split("\t");
            if (fields[0] === "error") return null;
            if (fields.length >= 2) states[fields[0]] = fields[1];
        }
        return states;
    }

    /**
     * メニュー項目の ✓ をオン・オフに揃えるよう、補助アプリに依頼する。
     * 切り替えはこのスクリプトが終わってから行われる
     * @param {Array} requests - [メニューの道筋, "on"|"off"|"toggle"] の配列。道筋は「表示>スマートガイド」の形
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function setAiMenuState(requests) {
        return launchAiMenuStateHelper(requests);
    }

    // メニューの ✓ を読む・揃える（再利用パーツ）ここまで / End of the reusable menu checkmark helpers

    // 設定キーで読めるトグル（再利用パーツ） / Menu toggles backed by preference keys (reusable)

    /**
     * 環境設定キーで状態を読めるメニューのトグルの一覧を返す
     * キーはメニューの切り替えにすぐ追従する（2026-10-09 実測）。✓ の表示は少し遅れて更新される
     * @returns {Object} 名前をキーに { pref: 環境設定キー, command: メニューコマンド ID } を持つオブジェクト
     */
    function getPrefMenuToggleTable() {
        return {
            glyphSnap:   { pref: "snapToGlyph",           command: "glyphSnapping" },               /* 表示 > グリフにスナップ / View > Snap to Glyph */
            pointSnap:   { pref: "snapToPoint",           command: "snappoint" },                   /* 表示 > ポイントにスナップ / View > Snap to Point */
            smartGuides: { pref: "smartGuides/isEnabled", command: "Snapomatic on-off menu item" }  /* 表示 > スマートガイド / View > Smart Guides */
        };
    }

    /**
     * メニューのトグルが入っているかを環境設定キーで読む
     * @param {string} toggleName - "glyphSnap" / "pointSnap" / "smartGuides"
     * @returns {boolean} 入っていれば true
     */
    function getPrefMenuToggle(toggleName) {
        return app.preferences.getBooleanPreference(getPrefMenuToggleTable()[toggleName].pref);
    }

    /**
     * メニューのトグルを指定の状態に揃える。異なるときだけ executeMenuCommand() で反転する
     * @param {string} toggleName - "glyphSnap" / "pointSnap" / "smartGuides"
     * @param {boolean} turnOn - 入れるなら true、切るなら false
     * @returns {boolean} 切り替えたら true、もともと指定の状態なら false
     */
    function setPrefMenuToggle(toggleName, turnOn) {
        if (getPrefMenuToggle(toggleName) === turnOn) return false;
        app.executeMenuCommand(getPrefMenuToggleTable()[toggleName].command);
        return true;
    }

    // 設定キーで読めるトグル（再利用パーツ）ここまで / End of the reusable preference-backed toggles

    // =========================================
    // 動作確認 / Demo
    // =========================================
    /* グリフにスナップを入れる（補助アプリ不要） / Turn on Snap to Glyph without the helper */
    setPrefMenuToggle("glyphSnap", true);

    var demoPaths = ["ウィンドウ>コントロール", "表示>グリッドにスナップ", "表示>ピクセルにスナップ", "表示>スマートガイド"];
    var demoStates = getAiMenuState(demoPaths);
    if (!demoStates) {
        alert("メニューの状態を読めませんでした。\n/Applications/SetAiMenuState.app とアクセシビリティの許可を確認してください。");
        return;
    }
    var report = [];
    for (var k = 0; k < demoPaths.length; k++) report.push(demoPaths[k] + " : " + demoStates[demoPaths[k]]);
    alert(report.join("\n"));

})();
