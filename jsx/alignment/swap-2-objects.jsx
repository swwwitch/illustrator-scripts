#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.swap-2-objects
Meta-Default-language=ja
Name=2つのオブジェクトの中心位置を入れ替える
Name[en]=Swap the center positions of two objects
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.3
Release-Date=2025-08-02
Tags=alignment,整列
X-com.swwwitch-Script-Name=swap-2-objects
X-com.swwwitch-Updated=2026-10-04
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/swap-2-objects.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/swap-2-objects.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/swap-2-objects.jsx
Description-BEGIN
[ja]
2つのオブジェクトを選択しているとき、それぞれの中心位置を入れ替えます。
通常オブジェクトは visibleBounds、クリップグループはマスクパスの geometricBounds を基準にします。
[en]
Swaps the center positions of two selected objects.
Ordinary objects use their visibleBounds, while clipping groups use the geometricBounds of the mask path.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

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
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: { ja: "2つのオブジェクトを選択してください。", en: "Select two objects." }
        }
    };

    // =========================================
    // 入れ替え / Swapping
    // =========================================

    /**
     * オブジェクトの中心座標を返す
     *
     * クリップグループはマスクパスの geometricBounds、それ以外は visibleBounds を基準にする。
     * マスクパスが見つからないクリップグループは visibleBounds にフォールバックする。
     * @param {PageItem} targetItem - 中心を求める対象のオブジェクト
     * @returns {number[]} 中心の [x, y] 座標
     */
    function getCenterPoint(targetItem) {
        var referenceBounds = targetItem.visibleBounds;

        if (targetItem.typename === "GroupItem" && targetItem.clipped) {
            /* クリップグループはマスクパスの範囲を基準にする / clipping groups follow the mask path */
            for (var i = 0; i < targetItem.pageItems.length; i++) {
                if (targetItem.pageItems[i].clipping) {
                    referenceBounds = targetItem.pageItems[i].geometricBounds;
                    break;
                }
            }
        }

        return [(referenceBounds[0] + referenceBounds[2]) / 2, (referenceBounds[1] + referenceBounds[3]) / 2];
    }

    /**
     * 2つのオブジェクトの中心位置を入れ替える
     * @param {PageItem} firstObject - 入れ替える一方のオブジェクト
     * @param {PageItem} secondObject - 入れ替えるもう一方のオブジェクト
     * @returns {void}
     */
    function swapObjectsByCenter(firstObject, secondObject) {
        var firstCenter = getCenterPoint(firstObject);
        var secondCenter = getCenterPoint(secondObject);

        /* 片方の移動量を求め、もう片方はその逆向きに動かす / one offset, applied in both directions */
        var dx = secondCenter[0] - firstCenter[0];
        var dy = secondCenter[1] - firstCenter[1];

        firstObject.translate(dx, dy);
        secondObject.translate(-dx, -dy);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を検証し、選択した2つのオブジェクトの中心位置を入れ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は TextRange が返り、length は文字数になる / A text selection returns a TextRange whose length counts characters */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 2) {
            alert(getLabel(LABELS.alert.selectTwo));
            return;
        }

        swapObjectsByCenter(selectedObjects[0], selectedObjects[1]);
    }

    main();

})();
