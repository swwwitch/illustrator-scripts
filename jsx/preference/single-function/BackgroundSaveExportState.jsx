#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.BackgroundSaveExportState
Name=BackgroundSaveExportState
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-18
Description-BEGIN
環境設定の［ファイル管理］→［バックグラウンドで保存］［バックグラウンドで書き出し］のオン／オフを読み出して表示します。
app.preferences では読めない（存在しないキー扱いで常に true が返る）ため、設定ファイル「Adobe Illustrator Cloud Prefs」を直接読みます。

Shows whether File Handling > Save in Background and Export in Background are on or off.
app.preferences cannot read them (the keys always return true), so this reads the "Adobe Illustrator Cloud Prefs" file directly.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

環境設定の［ファイル管理］→［バックグラウンドで保存］［バックグラウンドで書き出し］のオン／オフを読み出して表示します。
app.preferences では読めない（存在しないキー扱いで常に true が返る）ため、設定ファイル「Adobe Illustrator Cloud Prefs」を直接読みます。

### Overview

Shows whether File Handling > Save in Background and Export in Background are on or off.
app.preferences cannot read them (the keys always return true), so this reads the "Adobe Illustrator Cloud Prefs" file directly.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "BackgroundSaveExportState";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-18";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 設定ファイル / Preference file
    // =========================================

    /* 値の入っているファイル。メインの「Adobe Illustrator Prefs（環境設定）」にも同名キーがあるが、
       app.preferences の get で作られた残骸で、画面の設定とは連動しない
       The file holding the values. The main Prefs file also has keys with these names,
       but they are leftovers created by app.preferences getters and do not follow the UI */
    var CLOUD_PREFS_FILE_NAME = "Adobe Illustrator Cloud Prefs";

    /* 読み出す項目（key は設定ファイルのキー。大文字始まりの同名キーは別物なので区別する）
       Items to read (keys are case-sensitive; the capitalized variants are different keys) */
    var STATE_ITEMS = [
        { key: "enableBackgroundSave", labelKey: "save" },
        { key: "enableBackgroundExport", labelKey: "export" }
    ];

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
        item: {
            save: { ja: "バックグラウンドで保存", en: "Save in Background" },
            "export": { ja: "バックグラウンドで書き出し", en: "Export in Background" }
        },
        state: {
            on: { ja: "オン", en: "On" },
            off: { ja: "オフ", en: "Off" },
            unknown: { ja: "不明（キーが見つかりません）", en: "Unknown (key not found)" }
        },
        alert: {
            title: { ja: "ファイル管理", en: "File Handling" },
            fileNote: {
                ja: "起動中に変更した内容は、設定ファイルに書き込まれるまで反映されないことがあります。",
                en: "Changes made in this session may not appear until Illustrator writes the preference file."
            },
            fileNotFound: { ja: "設定ファイルが見つかりません：", en: "Preference file not found:" },
            readFailed: { ja: "設定ファイルを読めませんでした：", en: "Could not read the preference file:" }
        }
    };

    /**
     * { ja, en } から現在の言語の文言を返す
     * @param {Object} entry - { ja, en }
     * @returns {string} 文言
     */
    function getLabel(entry) {
        return entry ? (entry[uiLang] || entry.en || "") : "";
    }

    // =========================================
    // 処理 / Logic
    // =========================================

    /**
     * Cloud Prefs のファイルを返す。app.preferencesFolder が使えなければ、
     * 「Adobe Illustrator <メジャー番号> Settings」の下のロケールのフォルダーを探す
     * @returns {File|null} 見つからなければ null
     */
    function findCloudPrefsFile() {
        var candidateFolders = [];
        try {
            if (app.preferencesFolder) candidateFolders.push(app.preferencesFolder);
        } catch (e) {}

        var majorVersion = String(app.version).split(".")[0];
        var settingsFolder = new Folder(Folder.userData.parent.fsName + "/Preferences/Adobe Illustrator " + majorVersion + " Settings");
        if ($.os.indexOf("Windows") === 0) {
            settingsFolder = new Folder(Folder.userData.fsName + "/Adobe/Adobe Illustrator " + majorVersion + " Settings");
        }
        if (settingsFolder.exists) {
            candidateFolders.push(new Folder(settingsFolder.fsName + "/" + app.locale));
            var localeFolders = settingsFolder.getFiles(function (item) { return item instanceof Folder; });
            for (var i = 0; i < localeFolders.length; i++) candidateFolders.push(localeFolders[i]);
        }

        for (var j = 0; j < candidateFolders.length; j++) {
            var prefsFile = new File(candidateFolders[j].fsName + "/" + CLOUD_PREFS_FILE_NAME);
            if (prefsFile.exists) return prefsFile;
        }
        return null;
    }

    /**
     * ファイルを文字列で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 読めなければ null
     */
    function readText(textFile) {
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) return null;
        var text = textFile.read();
        textFile.close();
        return text;
    }

    /**
     * 設定ファイルの文字列から、行頭にある「/キー 値」の値を返す（大文字小文字は区別する）
     * @param {string} prefsText - 設定ファイルの中身（改行は CR）
     * @param {string} key - キー名
     * @returns {number|null} 値。キーが無ければ null
     */
    function readTopLevelInteger(prefsText, key) {
        var match = new RegExp("(?:^|[\\r\\n])/" + key + " (-?\\d+)").exec(prefsText);
        return match ? parseInt(match[1], 10) : null;
    }

    /**
     * 値を「オン／オフ／不明」の文言にする
     * @param {number|null} value - 設定値
     * @returns {string} 表示用の文言
     */
    function stateText(value) {
        if (value === null) return getLabel(LABELS.state.unknown);
        return getLabel(value ? LABELS.state.on : LABELS.state.off);
    }

    /**
     * 状態を読み出して alert で表示する
     */
    function main() {
        var prefsFile = findCloudPrefsFile();
        if (!prefsFile) {
            alert(getLabel(LABELS.alert.fileNotFound) + "\n" + CLOUD_PREFS_FILE_NAME);
            return;
        }
        var prefsText = readText(prefsFile);
        if (prefsText === null) {
            alert(getLabel(LABELS.alert.readFailed) + "\n" + prefsFile.fsName);
            return;
        }

        var lines = [];
        for (var i = 0; i < STATE_ITEMS.length; i++) {
            var item = STATE_ITEMS[i];
            lines.push(getLabel(LABELS.item[item.labelKey]) + (uiLang === "ja" ? " : " : ": ") + stateText(readTopLevelInteger(prefsText, item.key)));
        }
        lines.push("");
        lines.push(getLabel(LABELS.alert.fileNote));
        alert(lines.join("\n"), getLabel(LABELS.alert.title));
    }

    main();

})();
