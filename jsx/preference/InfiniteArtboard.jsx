#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

トリミング表示・アートボード・ガイド・ビデオ定規を隠し、カンバスカラーをホワイトにして、アートボードの境目が見えない「無限のカンバス」の表示にします。
すでに隠れている項目はそのままにします（メニューの反転と違い、表示に戻ってしまうことはありません）。

### 注意

メニューの状態を読んで切り替えるため、補助アプリ /Applications/SetAiMenuState.app（ai-scripts の helpers/SetAiMenuState.applescript）とアクセシビリティの許可が要ります。
メニューの切り替えはスクリプトの終了直後に行われます。

### Overview

Hides Trim View, artboards, guides and the video ruler, and sets the canvas color to white, so the canvas looks like one endless sheet.
Items already hidden stay hidden (unlike a plain menu toggle, nothing is turned back on).

### Notes

The menu states are read by the helper app /Applications/SetAiMenuState.app (helpers/SetAiMenuState.applescript in ai-scripts), which needs Accessibility permission.
The menus are switched right after the script ends.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InfiniteArtboard";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-09";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /*
     * オフにする（隠す）メニュー項目。名前が入れ替わる項目は「オフのときの名前|オンのときの名前」
     * 英語版のメニュー名は未検証
     * Menu items to turn off (hide). A label that flips is written "name while off|name while on".
     * The English menu names are unverified
     */
    var MENU_ITEMS_TO_HIDE = [
        { ja: "表示>トリミング表示", en: "View>Trim View" },
        { ja: "表示>アートボードを表示|アートボードを隠す", en: "View>Show Artboards|Hide Artboards" },
        { ja: "表示>ガイド>ガイドを表示|ガイドを隠す", en: "View>Guides>Show Guides|Hide Guides" },
        { ja: "表示>定規>ビデオ定規を表示|ビデオ定規を隠す", en: "View>Rulers>Show Video Rulers|Hide Video Rulers" }
    ];

    /* カンバスカラー（uiCanvasIsWhite）: 0＝UIに合わせる、1＝ホワイト / Canvas color: 0 = Match Brightness, 1 = White */
    var CANVAS_COLOR_WHITE = 1;

    // =========================================
    // メニューの状態 / Menu states
    // =========================================

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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            helperUnavailable: {
                ja: "カンバスカラーだけをホワイトにしました。\nトリミング表示・アートボード・ガイド・ビデオ定規を隠すには、/Applications/SetAiMenuState.app が必要です。",
                en: "Only the canvas color was set to white.\nHiding Trim View, artboards, guides and the video ruler needs /Applications/SetAiMenuState.app."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * カンバスカラーをホワイトにする。設定キーを書いただけでは描き直されないので、ズームを往復させる
     * @returns {void}
     */
    function setCanvasColorWhite() {
        app.preferences.setIntegerPreference("uiCanvasIsWhite", CANVAS_COLOR_WHITE);
        app.executeMenuCommand("zoomout");
        app.executeMenuCommand("zoomin");
    }

    /**
     * メイン処理：カンバスカラーをホワイトにし、メニュー項目を隠すよう補助アプリに依頼する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        setCanvasColorWhite();

        var hideRequests = [];
        for (var i = 0; i < MENU_ITEMS_TO_HIDE.length; i++) {
            hideRequests.push([MENU_ITEMS_TO_HIDE[i][uiLang] || MENU_ITEMS_TO_HIDE[i].en, "off"]);
        }
        /* 切り替えはこのスクリプトが終わった直後に行われる / Switched right after this script ends */
        if (!setAiMenuState(hideRequests)) alert(getLabel("alert.helperUnavailable"));
        app.redraw();
    }

    main();

})();
