#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式と文字ツメを、先頭の変数スイッチで
まとめて設定するスクリプトです（ダイアログは表示しない）。

- KERNING_MODE で自動カーニング方式を切り替える
  - 0 = なし／1 = メトリクス／2 = オプティカル／3 = 和文等幅
- TSUME_PERCENT で文字ツメを 0〜100% で設定する（デフォルト 0）
- テキストオブジェクト全体（選択範囲全体）に即適用する
- 「メトリクス」(1) のときだけプロポーショナルメトリクスをON、それ以外はOFF
- 未選択時・不正な KERNING_MODE 時はアラートを表示して何もしない

### Overview

Sets the auto-kerning method and Tsume (character aki) for the selected text
via switches at the top (no dialog is shown).

- KERNING_MODE switches the auto-kerning method
  - 0 = None / 1 = Metrics / 2 = Optical / 3 = Metrics - Roman Only
- TSUME_PERCENT sets Tsume from 0 to 100% (default 0)
- Applies immediately to the whole text object (entire selection)
- Proportional metrics is turned ON only for "Metrics" (1), OFF otherwise
- Shows an alert and stops when nothing is selected or KERNING_MODE is invalid

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================
    var SCRIPT_VERSION = "v1.0";

    // =========================================
    // 設定 / Configuration
    // =========================================

    /* 自動カーニング方式の切替スイッチ（AutoKernType の値）/ Auto-kerning method switch (AutoKernType values)
       0 = NOAUTOKERN       0           / None
       1 = AUTO             メトリクス   / Metrics
       2 = OPTICAL          オプティカル / Optical
       3 = METRICSROMANONLY 和文等幅     / Metrics - Roman Only */
    var KERNING_MODE = 2;

    /* 文字ツメの百分率 / Tsume percentage
       0〜100 の整数（1未満は0に切り捨て）。30 = 30%
       Integer 0-100 (values below 1 are truncated to 0). 30 = 30% */
    var TSUME_PERCENT = 0;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" },
            invalidMode: { ja: "KERNING_MODE の値が不正です", en: "Invalid KERNING_MODE value" }
        }
    };

    /* 言語に応じたラベル文字列を取得 / Resolve a label string for the current language */
    function getLocalizedText(entry) {
        if (!entry) return "";
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================

    /* 型名を安全に取得（host オブジェクトは typename を優先、JS オブジェクトは constructor.name）
       Safely resolve a type name: typename for host objects, constructor.name for JS objects */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            var itemType = getTypeName(selectedItem);
            if (itemType === "TextFrame") {
                selectedRanges.push(selectedItem.textRange);
            } else if (itemType === "TextRange") {
                selectedRanges.push(selectedItem);
            }
        }
        return selectedRanges;
    }

    // =========================================
    // カーニング / 文字ツメ処理 / Kerning & Tsume
    // =========================================

    /* KERNING_MODE（数値）を AutoKernType 列挙値に変換 / Map KERNING_MODE (number) to an AutoKernType enum
       enum プロパティには整数ではなく列挙値そのものを代入する必要がある
       The enum property must be assigned the enumerator itself, not a raw integer
       不正な値は null を返す / Returns null for an invalid value */
    function resolveKerningMethod(mode) {
        if (mode === 0) return AutoKernType.NOAUTOKERN;
        if (mode === 1) return AutoKernType.AUTO;
        if (mode === 2) return AutoKernType.OPTICAL;
        if (mode === 3) return AutoKernType.METRICSROMANONLY;
        return null;
    }

    /* 選択範囲にカーニング方式と文字ツメをまとめて適用 / Apply kerning method and Tsume to the given ranges
       メトリクスのときのみプロポーショナルメトリクスをON、それ以外はOFF
       Proportional metrics ON only for Metrics, OFF otherwise
       Tsume は 0〜100 の百分率（30 = 30%）/ Tsume is a percentage 0-100 (30 = 30%) */
    function applyAttributesToRanges(ranges, kerningMethod, tsumePercent) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < ranges.length; i++) {
            try {
                var attrs = ranges[i].characterAttributes;
                attrs.kerningMethod = kerningMethod;
                attrs.proportionalMetrics = useProportionalMetrics;
                attrs.Tsume = tsumePercent;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var kerningMethod = resolveKerningMethod(KERNING_MODE);
        if (kerningMethod === null) {
            alert(getLocalizedText(LABELS.alert.invalidMode));
            return;
        }

        var targetRanges = getSelectedTextRanges();
        if (targetRanges.length === 0) {
            alert(getLocalizedText(LABELS.alert.selectText));
            return;
        }

        applyAttributesToRanges(targetRanges, kerningMethod, TSUME_PERCENT);
        app.redraw();
    }

    main();

})();