#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式を「メトリクス」に設定します。
この方式のときのみプロポーショナルメトリクスをONにし、文字ツメは0%にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning-Metrics.md

### Overview

Sets the auto-kerning method of the selected text to Metrics.
Proportional metrics are turned on only for this method, and tsume is set to 0%.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning-Metrics.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerning-Metrics";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning-Metrics.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning-Metrics.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var KERNING_METHOD = AutoKernType.AUTO; /* 自動カーニング方式 / auto-kerning method */
    var TSUME_PERCENT  = 0; /* 文字ツメ（%）/ tsume (%) */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のテキスト範囲を集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextRange[]} テキスト範囲の配列
     */
    function getSelectedTextRanges(doc) {
        var currentSelection = doc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a single TextRange */
        if (currentSelection.typename === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem.typename === "TextFrame") {
                selectedRanges.push(selectedItem.textRange);
            } else if (selectedItem.typename === "TextRange") {
                selectedRanges.push(selectedItem);
            }
        }
        return selectedRanges;
    }

    /**
     * テキスト範囲にカーニング方式と文字ツメを適用する（プロポーショナルメトリクスはメトリクスのときだけ ON）
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @param {AutoKernType} kerningMethod - 自動カーニング方式
     * @param {number} tsumePercent - 文字ツメ（%）
     * @returns {void}
     */
    function applyKerningToRanges(textRanges, kerningMethod, tsumePercent) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < textRanges.length; i++) {
            /* 適用できない範囲は飛ばして続行 / Skip ranges that reject the attributes */
            try {
                var characterAttributes = textRanges[i].characterAttributes;
                characterAttributes.kerningMethod = kerningMethod;
                characterAttributes.proportionalMetrics = useProportionalMetrics;
                characterAttributes.Tsume = tsumePercent;
            } catch (e) {}
        }
    }

    /**
     * 選択中のテキストに設定を適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var targetRanges = getSelectedTextRanges(app.activeDocument);
        if (targetRanges.length === 0) return;
        applyKerningToRanges(targetRanges, KERNING_METHOD, TSUME_PERCENT);
    }

    main();

})();
