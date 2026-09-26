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

AiMemoPalette / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManagerPalette /
UnifiedTypePalette / ImportAndApplyGraphicStylePalette / ArtboardDisplayPresetManagerPalette /
TextCountStatsPalette / SelectionInspectorPalette / ApplyLeadingPerTextFramePalette / TextBreakSplitMergePalette /
AiAlignToArtboardPalette / AiSmartRotateViewPalette / AutoKerningPalette / FontPresetPickerPalette / KPTSketchyPalette /
LockHistoryPalette / PathInspectorPalette / QuickTransformPalette / TypeBasicsPalette /
ArtboardNavigatorPalette / LEConvertToShapePalette / AiSmartPathfinderPalette / SmartDistributorPalette /
AiAdjustVerticalGapPalette / DirectPrefsPalette / DocumentFontListSelectorPalette

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloseAllPalettes";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

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
    { name: "AiMemoPalette", engine: "TextMemoEngine", global: "__TextMemoWindow" },
    { name: "AiQuickPrefsPalette", engine: "AiQuickPrefsPalette", global: "__aiQuickPrefsPalette" },
    { name: "AiTextOutlineRestorePalette", engine: "TextOutlineWithMemo", global: "__textOutlineMemoPalette" },
    { name: "LinkedImageManagerPalette", engine: "LinkedImageManager", global: "__LIM_paletteWindow" },
    { name: "UnifiedTypePalette", engine: "UnifiedTypePanelEngine", global: "__UnifiedTypePanel" },
    { name: "ImportAndApplyGraphicStylePalette", engine: "ImportAndApplyGraphicStyle", global: "__importAndApplyGraphicStylePalette" },
    { name: "ArtboardDisplayPresetManagerPalette", engine: "ArtboardDisplayPresetManagerPalette", global: "__artboardDisplayPresetPalette" },
    { name: "TextCountStatsPalette", engine: "TextCountStatsSession", global: "__TextCountStatsPalette" },
    { name: "SelectionInspectorPalette", engine: "SelectionInspectorSession", global: "__SelectionInspectorPalette" },
    { name: "ApplyLeadingPerTextFramePalette", engine: "ApplyLeadingPerTextFrame", global: "__ALPTF_PALETTE__" },
    { name: "TextBreakSplitMergePalette", engine: "TextBreakSplitMergeEngine", global: "__TextBreakSplitMergePalette" },
    { name: "AiAlignToArtboardPalette", engine: "AiAlignToArtboard", global: "__aiAlignToArtboardWindow" },
    { name: "AiSmartRotateViewPalette", engine: "AiSmartRotateView", global: "__aiSmartRotateViewPalette" },
    { name: "AutoKerningPalette", engine: "AutoKerningPanelEngine", global: "__AutoKerningPanel" },
    { name: "FontPresetPickerPalette", engine: "FontPresetPickerEngine", global: "__FontPresetPicker" },
    { name: "KPTSketchyPalette", engine: "KPTSketchy", global: "__KPTSketchyPaletteWindow" },
    { name: "LockHistoryPalette", engine: "LockHistoryPalette", global: "__LockHistoryPaletteWindow" },
    { name: "PathInspectorPalette", engine: "PathInspectorSession", global: "__PathInspectorPalette" },
    { name: "QuickTransformPalette", engine: "QuickTransformPalette", global: "__quickTransformPalette" },
    { name: "TypeBasicsPalette", engine: "TypeBasicsPanelEngine", global: "__typeBasicsPanelInstance" },
    { name: "ArtboardNavigatorPalette", engine: "artboardNavigatorPalette", global: "artboardNavigatorWindow" },
    { name: "LEConvertToShapePalette", engine: "fxConvertToShape", global: "__fxConvertToShapePalette" },
    { name: "AiSmartPathfinderPalette", engine: "pathfinder-palette", global: "__pfPaletteWindow" },
    { name: "SmartDistributorPalette", engine: "smartDistributorPalette", global: "smartDistributorWindow" },
    { name: "AiAdjustVerticalGapPalette", engine: "AdjustVerticalGap", global: "__aiAdjustVerticalGapPalette" },
    { name: "DirectPrefsPalette", engine: "DirectPrefs", global: "__directPrefsPalette" },
    { name: "DocumentFontListSelectorPalette", engine: "DocumentFontListEngine", global: "__documentFontListSelectorPalette" }
];

// =========================================
// ローカライズ / Localization
// =========================================
/**
 * UI の表示言語を判定する（ロケールが ja 始まりなら日本語）
 * @returns {string} "ja" または "en"
 */
function detectUILanguage() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var uiLang = detectUILanguage();

var LABELS = {
    alert: {
        closedSome: { ja: "個のパレットを閉じました:", en: " palette(s) closed:" }
    }
};

/**
 * ラベルから表示言語の文言を返す
 * @param {Object} labelEntry - ja / en を持つラベル
 * @returns {string} 表示言語の文言（無ければ英語、それも無ければ空文字）
 */
function getLabel(labelEntry) {
    if (!labelEntry) return "";
    return labelEntry[uiLang] || labelEntry.en || "";
}

(function () {

    var closedNames = [];         // 実際に閉じたパレット名 / Names of palettes actually closed
    var pendingBridgeTalks = [];  // 応答が返るまで参照を保持（GC 防止）/ Keep references until responses return (prevent GC)

    /**
     * 閉じた結果をまとめて表示する（1 つも閉じなかったときは何も表示しない）
     * @returns {void}
     */
    function showSummary() {
        if (closedNames.length === 0) return;
        alert(closedNames.length + getLabel(LABELS.alert.closedSome) + "\n\n" + closedNames.join("\n"));
    }

    /**
     * 指定エンジンでパレット参照を閉じる BridgeTalk 本文を組み立てる。
     * 戻り値マーカー："CLOSED"＝閉じた / "IDLE"＝参照はあるが非表示 / "NONE"＝参照なし / "ERR"＝例外
     * @param {string} engineName - 受信側の常駐エンジン名
     * @param {string} globalName - $.global 上のパレット参照名
     * @returns {string} BridgeTalk 本文（先頭に #targetengine ディレクティブ）
     */
    function buildCloseBody(engineName, globalName) {
        return '#targetengine "' + engineName + '"\n' +
            '(function () {' +
            '    try {' +
            '        var paletteRef = $.global.' + globalName + ';' +
            '        if (!paletteRef) return "NONE";' +
            /* Window を直接保持する形式と { window: Window } のラッパー形式の両対応 / Accept both a bare Window and a { window: Window } wrapper */
            '        var paletteWindow = (paletteRef.close ? paletteRef : (paletteRef.window ? paletteRef.window : null));' +
            '        var wasOpen = false;' +
            '        try { if (paletteWindow && paletteWindow.visible) { wasOpen = true; paletteWindow.close(); } } catch (eClose) {}' +
            '        $.global.' + globalName + ' = null;' +
            '        return wasOpen ? "CLOSED" : "IDLE";' +
            '    } catch (e) { return "ERR"; }' +
            '})();';
    }

    /**
     * 1 通送って応答が返るまで BridgeTalk.pump() で同期的に待ち、閉じたら名前を記録する
     * @param {string} engineName - 受信側の常駐エンジン名
     * @param {string} globalName - $.global 上のパレット参照名
     * @param {string} paletteName - サマリー表示用の名前
     * @returns {void}
     */
    function closePaletteAndWait(engineName, globalName, paletteName) {
        var responseReceived = false;
        var resultMarker = "TIMEOUT";
        var closeRequest = new BridgeTalk();
        closeRequest.target = 'illustrator';
        closeRequest.body = buildCloseBody(engineName, globalName);
        closeRequest.onResult = function (response) { resultMarker = response.body; responseReceived = true; };
        closeRequest.onError = function () { resultMarker = "ERR"; responseReceived = true; };
        pendingBridgeTalks.push(closeRequest); // 応答が返るまで保持 / Retain until the response returns
        closeRequest.send();
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
        var paletteEntry = PALETTES[i];
        closePaletteAndWait(paletteEntry.engine, paletteEntry.global, paletteEntry.name);
    }

    if (SHOW_SUMMARY) showSummary();

})();
