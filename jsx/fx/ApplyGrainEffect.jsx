#target illustrator
#targetengine "ApplyGrainEffectEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyGrainEffect
Name=ApplyGrainEffect
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-12
Description-BEGIN
選択したオブジェクトに［効果］＞［テクスチャ］＞［粒状］を適用します。
パラメーターは Illustrator 標準のダイアログボックスで指定します。

Applies Effect > Texture > Grain to the selected objects.
Set the parameters in Illustrator's standard dialog box.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトに［効果］＞［テクスチャ］＞［粒状］を適用します。
パラメーターは Illustrator 標準のダイアログボックスで指定します。

### Overview

Applies Effect > Texture > Grain to the selected objects.
Set the parameters in Illustrator's standard dialog box.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyGrainEffect";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-12";                   /* 更新日 / last updated */

var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ApplyGrainEffect.jsx"; /* 最新版のソース / latest source */

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
