#target illustrator
#targetengine "RotateRomanInVerticalEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.RotateRomanInVertical
Name=RotateRomanInVertical
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-15
Meta-URL=https://github.com/swwwitch/illustrator-scripts
Description-BEGIN
選択したテキストに、文字パネルの［縦組み中の欧文回転］を適用します。
DOM では設定できないため、一時アクション（ダイナミックアクション）で実行します。

Applies "Rotate Standard Roman Characters in Vertical Text" from the Character panel to the selected text.
The DOM cannot set it, so it runs through a temporary (dynamic) action.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストに、文字パネルの［縦組み中の欧文回転］を適用します。
DOM では設定できないため、一時アクション（ダイナミックアクション）で実行します。

### Overview

Applies "Rotate Standard Roman Characters in Vertical Text" from the Character panel to the selected text.
The DOM cannot set it, so it runs through a temporary (dynamic) action.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RotateRomanInVertical";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-15";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

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
            noSelection: { ja: "テキストを選択してください。", en: "Select text first." },
            actionFailed: { ja: "縦組み中の欧文回転を適用できませんでした。", en: "Could not apply Rotate Standard Roman Characters in Vertical Text." }
        }
    };

    // =========================================
    // 一時アクション定義 / Action definition
    // =========================================
    var ACTION_SET_NAME = "RotateRomanInVertical";   /* アクションセット名 / action set name */
    var ACTION_NAME = "rotateRoman";                 /* アクション名 / action name */

    /**
     * 文字パネルの［縦組み中の欧文回転］（キー 'rriv'）だけをオンにするアクション定義を返す
     * （文字パネルのイベントは、書いたパラメーターだけを適用する）
     * @returns {string} アクション定義のテキスト
     */
    function buildRotateRomanAction() {
        return ["/version 3"].concat(
            buildActionNameLines("", ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount 1", "/action-1 {"],
            buildActionNameLines("\t", ACTION_NAME),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (adobe_SLOCharacterPalette)"
            ],
            buildActionNameLines("\t\t", "文字", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key 1920100726",
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (boolean)",
                "\t\t\t/value 1",
                "\t\t}",
                "\t}",
                "}",
                ""
            ]
        ).join("\n");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択がテキスト（文字選択、またはテキストオブジェクトを含む選択）かを返す
     * @param {*} selection - app.activeDocument.selection
     * @returns {boolean} テキストなら true
     */
    function isTextSelection(selection) {
        if (!selection) return false;
        /* 文字選択中の selection は TextRange で [0] が無い / A text selection is a TextRange without [0] */
        if (selection.typename === "TextRange") return true;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") return true;
        }
        return false;
    }

    /**
     * 選択したテキストに［縦組み中の欧文回転］を適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || !isTextSelection(app.activeDocument.selection)) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        if (!runTemporaryAction(buildRotateRomanAction(), ACTION_SET_NAME, ACTION_NAME)) {
            alert(getLabel("alert.actionFailed"));
        }
    }

    main();

})();
