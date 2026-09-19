#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ。

- 各パレットは `#targetengine` で個別の常駐エンジンに載っており、`$.global` はエンジン
  ごとに独立している。そのため 1 本のスクリプトから他エンジンのパレット参照を直接は
  参照できない
- そこで各エンジンごとに `#targetengine` 付きの BridgeTalk を 1 通ずつ送り、
  受信側エンジンで `$.global.<参照名>` を読んで開いていれば `close()` する
- パレット参照は「Window を直接保持」する形式と「{ window: Window } のラッパー」形式の
  両方に対応する
- 閉じた後は `$.global.<参照名>` を null にして参照を解放する（各パレット本体の onClose
  でも解放されるが保険）
- 対象は PALETTES テーブルで管理。パレットを増やしたら 1 行追加するだけで対象にできる

### 実行時の要点 / Runtime notes

- BridgeTalk のクロスエンジン往復は「ファイル＞スクリプト」実行でも同期的に効く。ただし
  応答が返るまで送信側スクリプトを生かしておく必要があるため、送信ごとに BridgeTalk.pump()
  で応答を待つ（待たずにスクリプトが終わると配信前に終了して閉じ損ねる）
- 送信した BridgeTalk オブジェクトは応答が返るまで配列に保持する（途中で GC されると
  メッセージが配信されず閉じ損ねる）
- 1 件ずつ順に閉じることで、複数エンジンの close() が同一 UI スレッド上で重なって
  ハングするのを防ぐ

### 対象パレット / Target palettes

AiMemoPallete / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManager /
UnifiedTypePanel / ImportAndApplyGraphicStyle / ArtboardDisplayPresetManager /
TextCountStats / SelectionInspector / ApplyLeadingPerTextFrame / TextBreakSplitMergePallete

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloseAllPalettes";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================
/* 閉じた結果をまとめてアラート表示するか / Whether to show a summary alert after closing
   true  = 結果を表示 / Show the result
   false = 何も表示しない（サイレント）/ Silent */
var SHOW_SUMMARY = true;

/* 1 パレットあたりの応答待ち上限（ミリ秒）。応答は通常数ミリ秒で返るので、これは無応答時の
   保険上限 / Max wait per palette (ms); responses normally return in a few ms, so this is just a
   safety cap for an unresponsive engine */
var MAX_WAIT_MS = 2000;

/* 応答待ちループの 1 回あたりの休止（ミリ秒）/ Sleep per pump iteration while waiting (ms) */
var POLL_INTERVAL_MS = 10;

// =========================================
// 対象パレット定義 / Target palette definitions
// =========================================
/* name   : 表示用の名前 / Display name
   engine : #targetengine で宣言された常駐エンジン名 / Persistent engine name declared via #targetengine
   global : そのエンジンの $.global に載るパレット参照名 / Palette reference name kept on that engine's $.global */
var PALETTES = [
    { name: "AiMemoPallete", engine: "TextMemoEngine", global: "__TextMemoWindow" },
    { name: "AiQuickPrefsPalette", engine: "AiQuickPrefsPalette", global: "__aiQuickPrefsPalette" },
    { name: "AiTextOutlineRestorePalette", engine: "TextOutlineWithMemo", global: "__textOutlineMemoPalette" },
    { name: "LinkedImageManager", engine: "LinkedImageManager", global: "__LIM_paletteWindow" },
    { name: "UnifiedTypePanel", engine: "UnifiedTypePanelEngine", global: "__UnifiedTypePanel" },
    { name: "ImportAndApplyGraphicStyle", engine: "ImportAndApplyGraphicStyle", global: "__importAndApplyGraphicStylePalette" },
    { name: "ArtboardDisplayPresetManager", engine: "PresetManagerArtboardsPalette", global: "__artboardPresetPalette" },
    { name: "TextCountStats", engine: "TextCountStatsSession", global: "__TextCountStatsPalette" },
    { name: "SelectionInspector", engine: "SelectionInspectorSession", global: "__SelectionInspectorPalette" },
    { name: "ApplyLeadingPerTextFrame", engine: "ApplyLeadingPerTextFrame", global: "__ALPTF_PALETTE__" },
    { name: "TextBreakSplitMergePallete", engine: "TextBreakSplitMergeEngine", global: "__TextBreakSplitMergePalette" }
];

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の言語を判定（ロケールが ja 始まりなら日本語）/ Detect UI language (Japanese if locale starts with "ja") */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    message: {
        closedSome: { ja: "個のパレットを閉じました:", en: " palette(s) closed:" },
        closedNone: { ja: "開いているパレットはありませんでした。", en: "No open palettes were found." }
    }
};

/* ラベルノードから現在言語の文言を返す / Resolve a label node to the current language */
function L(labelNode) {
    if (!labelNode) return "";
    return labelNode[currentLanguage] || labelNode.en || "";
}

(function () {

    var closedNames = [];   // 実際に閉じたパレット名 / Names of palettes actually closed
    var bridges = [];       // 応答が返るまで参照を保持（GC 防止）/ Keep references until responses return (prevent GC)

    /* 閉じた結果をまとめて表示 / Show a summary of what was closed */
    function showSummary() {
        if (closedNames.length === 0) {
            alert(L(LABELS.message.closedNone));
            return;
        }
        alert(closedNames.length + L(LABELS.message.closedSome) + "\n\n" + closedNames.join("\n"));
    }

    /* 指定エンジンでパレット参照を閉じる BridgeTalk 本文を組み立て /
       Build the BridgeTalk body that closes a palette reference in the given engine */
    /* @param {string} engineName - 受信側の常駐エンジン名 / Target persistent engine name
       @param {string} globalName - $.global 上のパレット参照名 / Palette reference name on $.global
       @returns {string} BridgeTalk 本文（先頭に #targetengine ディレクティブ）/ BridgeTalk body (prefixed with the #targetengine directive)
       戻り値マーカー / Return markers: "CLOSED"=閉じた / "IDLE"=参照はあるが非表示 / "NONE"=参照なし / "ERR"=例外 */
    function buildCloseBody(engineName, globalName) {
        return '#targetengine "' + engineName + '"\n' +
            '(function () {' +
            '    try {' +
            '        var paletteRef = $.global.' + globalName + ';' +
            '        if (!paletteRef) return "NONE";' +
            /* Window を直接保持する形式と { window: Window } のラッパー形式の両対応 */
            '        var paletteWindow = (paletteRef.close ? paletteRef : (paletteRef.window ? paletteRef.window : null));' +
            '        var wasOpen = false;' +
            '        try { if (paletteWindow && paletteWindow.visible) { wasOpen = true; paletteWindow.close(); } } catch (eClose) {}' +
            '        $.global.' + globalName + ' = null;' +
            '        return wasOpen ? "CLOSED" : "IDLE";' +
            '    } catch (e) { return "ERR"; }' +
            '})();';
    }

    /* 1 通送って応答が返るまで BridgeTalk.pump() で同期的に待ち、閉じたら名前を記録 /
       Send one message, block on BridgeTalk.pump() until the response returns, and record the name if closed */
    /* @param {string} engineName - 受信側の常駐エンジン名 / Target persistent engine name
       @param {string} globalName - $.global 上のパレット参照名 / Palette reference name on $.global
       @param {string} paletteName - サマリー表示用の名前 / Display name for the summary */
    function closeOneAndWait(engineName, globalName, paletteName) {
        var responseReceived = false;
        var resultMarker = "TIMEOUT";
        var bridge = new BridgeTalk();
        bridge.target = 'illustrator';
        bridge.body = buildCloseBody(engineName, globalName);
        bridge.onResult = function (response) { resultMarker = response.body; responseReceived = true; };
        bridge.onError = function () { resultMarker = "ERR"; responseReceived = true; };
        bridges.push(bridge); // 応答が返るまで保持 / Retain until the response returns
        bridge.send();
        var elapsedMs = 0;
        while (!responseReceived && elapsedMs < MAX_WAIT_MS) {
            BridgeTalk.pump(); // 保留中のメッセージを処理して onResult/onError を発火 / Process pending messages so onResult/onError fire
            $.sleep(POLL_INTERVAL_MS);
            elapsedMs += POLL_INTERVAL_MS;
        }
        if (resultMarker === 'CLOSED') closedNames.push(paletteName);
    }

    // =========================================
    // 各エンジンを 1 件ずつ確実に閉じる / Close each engine one at a time, reliably
    // =========================================
    for (var i = 0; i < PALETTES.length; i++) {
        var palette = PALETTES[i];
        closeOneAndWait(palette.engine, palette.global, palette.name);
    }

    if (SHOW_SUMMARY) showSummary();

})();
