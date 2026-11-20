#target illustrator
#targetengine "ApplyGrainEffectEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyGrainEffect
Meta-Default-language=ja
Name=［粒状］効果を適用
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-12
Tags=effect,効果
X-com.swwwitch-Script-Name=ApplyGrainEffect
X-com.swwwitch-Updated=2026-10-12
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ApplyGrainEffect.jsx
Description-BEGIN
[ja]
選択したオブジェクトに［効果］＞［テクスチャ］＞［粒状］を適用します。
パラメーターは Illustrator 標準のダイアログボックスで指定します。
[en]
Applies Effect > Texture > Grain to the selected objects.
Set the parameters in Illustrator's standard dialog box.
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
    // ユーザー設定 / User Settings
    // =========================================

    /* ［粒状］のメニューコマンド。末尾の空白も ID の一部 / Grain menu command; the trailing space is part of the ID */
    var GRAIN_MENU_COMMAND = "Live PSAdapter_plugin_Grn ";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    var LABELS = {
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:  { ja: "オブジェクトを選択してください。", en: "Select one or more objects." },
            textEditing:  { ja: "文字の編集中は実行できません。オブジェクトを選択してください。", en: "Cannot run while editing text. Select objects instead." },
            commandFailed: { ja: "［粒状］を適用できませんでした。", en: "Could not apply Grain." }
        }
    };

    /**
     * LABELS から今の UI 言語の文言を取り出す
     * @param {string} labelPath - "alert.noSelection" のようなパス
     * @returns {string} 文言。見つからなければパスの文字列
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelEntry = LABELS;
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
        return (labelEntry != null && labelEntry[uiLang] != null) ? labelEntry[uiLang] : labelPath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したオブジェクトに［粒状］を適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var docSelection = doc.selection;

        /* 文字の編集中は TextRange が入り、配列ではない / While editing text, selection is a TextRange, not an array */
        if (docSelection && docSelection.typename === "TextRange") {
            alert(getLabel("alert.textEditing"));
            return;
        }
        if (!docSelection || docSelection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        try {
            /* 標準のダイアログボックスが開き、選択中の各オブジェクトに適用される
               Opens the standard dialog box and applies the effect to each selected object */
            app.executeMenuCommand(GRAIN_MENU_COMMAND);
        } catch (e) {
            alert(getLabel("alert.commandFailed") + "\n" + e);
        }
    }

    main();
})();
