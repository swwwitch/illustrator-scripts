#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式を「オプティカル」に設定します。
プロポーショナルメトリクスはOFF、文字ツメは0%にします。

詳細は README を参照してください。

### Overview

Sets the auto-kerning method of the selected text to Optical.
Proportional metrics are turned off and tsume is set to 0%.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerning-Optical";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning-Optical.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning-Optical.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 型名を安全に取得（host オブジェクトは typename を優先、JS オブジェクトは constructor.name） */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 選択中のテキスト範囲を取得 */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる */
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

    /* 選択範囲にカーニング方式を適用（メトリクスのときのみプロポーショナルメトリクスをON、文字ツメは0%に） */
    function applyKerningToRanges(ranges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.kerningMethod = kerningMethod;
                ranges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
                ranges[i].characterAttributes.Tsume = 0;
            } catch (e) {
                // 適用できない範囲はスキップ
            }
        }
    }

    function main() {
        if (app.documents.length <= 0) return;
        var targetRanges = getSelectedTextRanges();
        if (targetRanges.length === 0) return;
        applyKerningToRanges(targetRanges, AutoKernType.OPTICAL);
    }

    main();

})();
