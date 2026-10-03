#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

常駐エンジンで動いている各種フローティングパレットをまとめて閉じるユーティリティ。

- `$.global` はエンジンごとに独立しており、Illustrator では外から別エンジンへコードを届ける
  手段が無い（BridgeTalk 本文の `#targetengine` は無視され main で動く）
- そこで対象パレットと同じ共通エンジン `SwwwitchPalettes` で動き、`$.global.<参照名>` を
  直接読んで開いていれば `close()` する
- 共通エンジンへ移行していないパレットは閉じられない（参照が見えない）
- パレット参照は「Window を直接保持」する形式と「{ window: Window } のラッパー」形式の
  両方に対応する
- 対象は PALETTES テーブルで管理

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloseAllPalettes.md

### 実行時の要点 / Runtime notes

- 共通エンジンでは IIFE の外の変数・関数がパレット同士で共有される。このスクリプトも
  基本情報ブロック以外は IIFE の中に置く

### 対象パレット / Target palettes

AiMemoPalette / AiQuickPrefsPalette / AiTextOutlineRestorePalette / LinkedImageManagerPalette /
UnifiedTypePalette / ImportAndApplyGraphicStylePalette / ArtboardDisplayPresetManagerPalette /
TextCountStatsPalette / SelectionInspectorPalette / ApplyLeadingPerTextFramePalette / TextProcessingPalette /
AiAlignToArtboardPalette / AiSmartRotateViewPalette / AutoKerningPalette / FontPresetPickerPalette / KPTSketchyPalette /
LockHistoryPalette / PathInspectorPalette / QuickTransformPalette / TypeBasicsPalette /
ArtboardNavigatorPalette / LEConvertToShapePalette / AiSmartPathfinderPalette / SmartDistributorPalette /
AiAdjustVerticalGapPalette / DirectPrefsPalette / DocumentFontListSelectorPalette

### Overview

Closes the floating palettes that run in persistent engines, all at once.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloseAllPalettes.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloseAllPalettes";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-03";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloseAllPalettes.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloseAllPalettes.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    /* 閉じた結果をまとめてアラート表示するか / Whether to show a summary alert after closing
       true  = 結果を表示 / Show the result
       false = 何も表示しない（サイレント）/ Silent */
    var SHOW_SUMMARY = true;

    // =========================================
    // 対象パレット定義 / Target palette definitions
    // =========================================
    /* name   : 表示用の名前 / Display name
       global : 共通エンジンの $.global に載るパレット参照名 / Palette reference name kept on the shared engine's $.global */
    var PALETTES = [
        { name: "AiMemoPalette", global: "__TextMemoWindow" },
        { name: "AiQuickPrefsPalette", global: "__aiQuickPrefsPalette" },
        { name: "AiTextOutlineRestorePalette", global: "__textOutlineMemoPalette" },
        { name: "LinkedImageManagerPalette", global: "__LIM_paletteWindow" },
        { name: "UnifiedTypePalette", global: "__UnifiedTypePanel" },
        { name: "ImportAndApplyGraphicStylePalette", global: "__importAndApplyGraphicStylePalette" },
        { name: "ArtboardDisplayPresetManagerPalette", global: "__artboardDisplayPresetPalette" },
        { name: "TextCountStatsPalette", global: "__TextCountStatsPalette" },
        { name: "SelectionInspectorPalette", global: "__SelectionInspectorPalette" },
        { name: "ApplyLeadingPerTextFramePalette", global: "__ALPTF_PALETTE__" },
        { name: "TextProcessingPalette", global: "__TextBreakSplitMergePalette" },
        { name: "AiAlignToArtboardPalette", global: "__aiAlignToArtboardWindow" },
        { name: "AiSmartRotateViewPalette", global: "__aiSmartRotateViewPalette" },
        { name: "AutoKerningPalette", global: "__AutoKerningPanel" },
        { name: "FontPresetPickerPalette", global: "__FontPresetPicker" },
        { name: "KPTSketchyPalette", global: "__KPTSketchyPaletteWindow" },
        { name: "LockHistoryPalette", global: "__LockHistoryPaletteWindow" },
        { name: "PathInspectorPalette", global: "__PathInspectorPalette" },
        { name: "QuickTransformPalette", global: "__quickTransformPalette" },
        { name: "TypeBasicsPalette", global: "__typeBasicsPanelInstance" },
        { name: "ArtboardNavigatorPalette", global: "artboardNavigatorWindow" },
        { name: "LEConvertToShapePalette", global: "__fxConvertToShapePalette" },
        { name: "AiSmartPathfinderPalette", global: "__pfPaletteWindow" },
        { name: "SmartDistributorPalette", global: "smartDistributorWindow" },
        { name: "AiAdjustVerticalGapPalette", global: "__aiAdjustVerticalGapPalette" },
        { name: "DirectPrefsPalette", global: "__directPrefsPalette" },
        { name: "DocumentFontListSelectorPalette", global: "__documentFontListSelectorPalette" },
        { name: "FavoriteFontPickerPalette", global: "__FavoriteFontPickerPalette" }
    ];

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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
            closedSome: { ja: "個のパレットを閉じました:", en: " palette(s) closed:" }
        }
    };

    var closedNames = []; // 実際に閉じたパレット名 / Names of palettes actually closed

    /**
     * 閉じた結果をまとめて表示する（1 つも閉じなかったときは何も表示しない）
     * @returns {void}
     */
    function showSummary() {
        if (closedNames.length === 0) return;
        alert(closedNames.length + getLabel(LABELS.alert.closedSome) + "\n\n" + closedNames.join("\n"));
    }

    /**
     * $.global のパレット参照を閉じて解放する
     * @param {string} globalName - $.global 上のパレット参照名
     * @returns {boolean} 開いていたパレットを閉じたら true
     */
    function closePalette(globalName) {
        var paletteRef = $.global[globalName];
        if (!paletteRef) return false;
        /* Window を直接保持する形式と { window: Window } のラッパー形式の両対応 / Accept both a bare Window and a { window: Window } wrapper */
        var paletteWindow = paletteRef.close ? paletteRef : paletteRef.window;
        var wasOpen = false;
        try {
            if (paletteWindow && paletteWindow.visible) {
                wasOpen = true;
                paletteWindow.close();
            }
        } catch (staleReferenceError) {}
        $.global[globalName] = null; // onClose でも解放されるが保険 / onClose also releases it; this is a safety net
        return wasOpen;
    }

    for (var i = 0; i < PALETTES.length; i++) {
        if (closePalette(PALETTES[i].global)) closedNames.push(PALETTES[i].name);
    }

    if (SHOW_SUMMARY) showSummary();

})();
