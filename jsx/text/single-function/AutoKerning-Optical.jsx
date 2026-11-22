#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AutoKerning-Optical
Meta-Default-language=ja
Name=自動カーニングを「オプティカル」に設定
Name[en]=Set auto-kerning to Optical
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0
Tags=text,テキスト
X-com.swwwitch-Script-Name=AutoKerning-Optical
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning-Optical.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning-Optical.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/AutoKerning-Optical.jsx
Description-BEGIN
[ja]
選択したテキストの自動カーニング方式を「オプティカル」に設定します。
プロポーショナルメトリクスはOFF、文字ツメは0%にします。
[en]
Sets the auto-kerning method of the selected text to Optical.
Proportional metrics are turned off and tsume is set to 0%.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var KERNING_METHOD = AutoKernType.OPTICAL; /* 自動カーニング方式 / auto-kerning method */
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
