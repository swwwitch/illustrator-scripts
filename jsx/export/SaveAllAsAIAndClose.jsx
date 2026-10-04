#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

開いているすべてのドキュメントを .ai 形式で元のフォルダーに保存し、閉じます。
未保存（保存先のない）ドキュメントはスキップします。

### Overview

Saves every open document as an .ai file in its original folder, then closes it.
Documents that have never been saved are skipped.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SaveAllAsAIAndClose";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var SAVE_EXTENSION     = ".ai"; /* 保存する拡張子 / Extension of the saved file */
    var PDF_COMPATIBLE     = true;  /* PDF 互換ファイルを作成 / Embed PDF-compatible data */
    var EMBED_LINKED_FILES = false; /* 配置画像を含む / Embed linked files */
    var EMBED_ICC_PROFILE  = false; /* ICC プロファイルを埋め込む / Embed the ICC profile */
    var USE_COMPRESSION    = true;  /* 圧縮を使用 / Use compression */

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

    var LABELS = {
        alert: {
            title: { ja: "すべて .ai で保存して閉じる", en: "Save All as .ai and Close" },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            saveFailed: {
                ja: "次のドキュメントは保存できなかったため、開いたままにしました。",
                en: "The following documents could not be saved and were left open."
            },
            skippedUnsaved: {
                ja: "次のドキュメントは一度も保存されていないため、スキップしました。",
                en: "The following documents have never been saved and were skipped."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 開いているドキュメントを後ろから順に .ai で保存して閉じる
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"), getLabel("alert.title"), true);
            return;
        }

        var aiSaveOptions = getAiSaveOptions();
        var failedLines = [];
        var skippedNames = [];
        var prevInteractionLevel = app.userInteractionLevel;
        try {
            /* 閉じるときの確認ダイアログを出さない / Suppress confirmation dialogs while closing */
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

            /* 後ろから処理して、閉じた分の添字のずれを避ける / Iterate backwards so closing a document does not shift the ones still to come */
            for (var i = app.documents.length - 1; i >= 0; i--) {
                var sourceDoc = app.documents[i];

                var destFolder = getSavedFolder(sourceDoc);
                if (!destFolder) {
                    skippedNames.push(sourceDoc.name);
                    continue;
                }

                var targetFile = getTargetFile(sourceDoc.name, SAVE_EXTENSION, destFolder);

                /* 失敗しても次のドキュメントへ進む / Move on to the next document even if this one fails */
                try {
                    sourceDoc.saveAs(targetFile, aiSaveOptions);
                } catch (e) {
                    failedLines.push(sourceDoc.name + " — " + e.message);
                    continue;
                }

                /* 保存直後なので変更は残っていない / Nothing is left unsaved right after saving */
                sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
            }

        } finally {
            /* 操作レベルを元に戻す / Restore the interaction level */
            app.userInteractionLevel = prevInteractionLevel;
        }

        showResultAlert(failedLines, skippedNames);
    }

    /**
     * 保存できなかったものとスキップしたものを1つのアラートにまとめて表示する（どちらも無ければ何もしない）
     * @param {string[]} failedLines - 「ドキュメント名 — エラー内容」の一覧
     * @param {string[]} skippedNames - 未保存でスキップしたドキュメント名の一覧
     * @returns {void}
     */
    function showResultAlert(failedLines, skippedNames) {
        var messageSections = [];
        if (failedLines.length > 0) {
            messageSections.push(getLabel("alert.saveFailed") + "\n" + failedLines.join("\n"));
        }
        if (skippedNames.length > 0) {
            messageSections.push(getLabel("alert.skippedUnsaved") + "\n" + skippedNames.join("\n"));
        }
        if (messageSections.length === 0) return;
        alert(messageSections.join("\n\n"), getLabel("alert.title"), true);
    }

    /**
     * ユーザー設定から .ai 保存オプションを作る
     * @returns {IllustratorSaveOptions} 保存オプション
     */
    function getAiSaveOptions() {
        var aiSaveOptions = new IllustratorSaveOptions();
        aiSaveOptions.pdfCompatible = PDF_COMPATIBLE;
        aiSaveOptions.embedLinkedFiles = EMBED_LINKED_FILES;
        aiSaveOptions.embedICCProfile = EMBED_ICC_PROFILE;
        aiSaveOptions.compressed = USE_COMPRESSION;
        /* aiSaveOptions.compatibility = Compatibility.ILLUSTRATOR17; // 互換バージョンの例 / Example of a compatibility version */
        return aiSaveOptions;
    }

    /**
     * 保存済みドキュメントのフォルダーを返す（未保存なら null）
     * @param {Document} sourceDoc - 対象のドキュメント
     * @returns {Folder|null} 保存先フォルダー。未保存なら null
     */
    function getSavedFolder(sourceDoc) {
        /* 未保存だと path は例外か空のフォルダーになる / An unsaved document throws on path or returns an empty folder */
        var docFolder;
        try {
            docFolder = sourceDoc.path;
        } catch (e) {
            return null;
        }
        if (!docFolder || String(docFolder.fsName) === "") return null;
        return docFolder;
    }

    /**
     * ドキュメント名の拡張子を差し替えた保存先ファイルを返す
     * @param {string} docName - ドキュメント名
     * @param {string} extension - 拡張子（例: ".ai"）
     * @param {Folder} destFolder - 保存先フォルダー
     * @returns {File} 保存先ファイル
     */
    function getTargetFile(docName, extension, destFolder) {
        /* 拡張子の差し替え / Replace the extension */
        var dotIndex = docName.lastIndexOf(".");
        var baseName = (dotIndex < 0) ? docName : docName.substring(0, dotIndex);
        return new File(destFolder + "/" + baseName + extension);
    }

    main();

})();
