#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字に自動サイズ調整を適用します（拡張のみ）。
FitAreaText の［自動サイズ調整］だけを、ダイアログなしで実行する版です。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoFitTextFrame.md

### Overview

Applies Auto Size to the selected area type (expand only).
A dialog-free version of the Auto Size option in FitAreaText.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoFitTextFrame.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoFitTextFrame";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoFitTextFrame.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoFitTextFrame.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            selectObject: {
                ja: "エリア内文字を選択してください。",
                en: "Please select area type."
            },
            noAreaText: {
                ja: "エリア内文字が選択されていません。\n自動サイズ調整はエリア内文字のみ対応です。",
                en: "No area type found in the selection.\nAuto Size is supported for area type only."
            },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 自動サイズ調整アクションの値（ON）/ Value of the Auto Size action (on) */
    var AUTO_SIZE_ON = 1;

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    /**
     * 処理対象にできるエリア内文字か判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 対象にできるとき true
     */
    function isAdjustableAreaText(textFrame) {
        return (textFrame.kind == TextType.AREATEXT &&
            textFrame.editable && !textFrame.locked && !textFrame.hidden);
    }

    /**
     * 選択項目を再帰的にたどってエリア内文字を集める
     * @param {Object} selectedItem - 選択項目（TextRange／TextFrame／GroupItem など）
     * @param {TextFrame[]} collectedFrames - 集めたエリア内文字の配列（破壊的に追加）
     * @returns {void}
     */
    function collectAreaTextFromItem(selectedItem, collectedFrames) {
        if (!selectedItem) return;

        /* 選択の種類によっては parent や pageItems を読めない / Some selections cannot expose parent or pageItems */
        try {
            /* 文字カーソルでの選択はフレームに読み替える / A TextRange selection is read as its frame */
            if (selectedItem.typename === "TextRange") {
                if (selectedItem.parent && selectedItem.parent.typename === "TextFrame") {
                    collectAreaTextFromItem(selectedItem.parent, collectedFrames);
                }
                return;
            }

            if (selectedItem.typename === "TextFrame") {
                if (isAdjustableAreaText(selectedItem)) collectedFrames.push(selectedItem);
                return;
            }

            /* グループ・レイヤーなどは中身をたどる / Containers are traversed */
            if (selectedItem.pageItems) {
                for (var i = 0; i < selectedItem.pageItems.length; i++) {
                    collectAreaTextFromItem(selectedItem.pageItems[i], collectedFrames);
                }
            }
        } catch (e) { }
    }

    /**
     * 選択から処理対象のエリア内文字を重複なく集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 処理対象のエリア内文字
     */
    function getSelectedAreaText(doc) {
        var collectedFrames = [];
        for (var i = 0; i < doc.selection.length; i++) {
            collectAreaTextFromItem(doc.selection[i], collectedFrames);
        }

        /* 参照そのもので重複を除く（文字列化では区別できない）/ De-duplicate by object reference */
        var uniqueFrames = [];
        for (var j = 0; j < collectedFrames.length; j++) {
            var isDuplicate = false;
            for (var k = 0; k < uniqueFrames.length; k++) {
                if (uniqueFrames[k] === collectedFrames[j]) {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate) uniqueFrames.push(collectedFrames[j]);
        }
        return uniqueFrames;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 自動サイズ調整 / Auto Size
    // =========================================

    /**
     * 自動サイズ調整をアクション経由で切り替える
     * @param {number} autoSizeValue - 1（ON）または 2（OFF）
     * @returns {boolean} 実行できたら true
     */
    function setAutoSizeByAction(autoSizeValue) {
        /* アクション定義（セット名 AreaType／アクション名 AutoSize）/ Action definition */
        var actionCode = [
            '/version 3',
            '/name [ 8 4172656154797065]',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '  /name [ 8 4175746f53697a65 ]',
            '  /keyIndex 0',
            '  /colorIndex 0',
            '  /isOpen 1',
            '  /eventCount 1',
            '  /event-1 {',
            '    /useRulersIn1stQuadrant 0',
            '    /internalName (adobe_SLOAreaTextDialog)',
            '    /localizedName [ 33',
            '      e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3',
            '    ]',
            '    /isOpen 1',
            '    /isOn 1',
            '    /hasDialog 0',
            '    /parameterCount 1',
            '    /parameter-1 {',
            '      /key 1952539754',
            '      /showInPalette 4294967295',
            '      /type (integer)',
            '      /value ' + String(autoSizeValue),
            '    }',
            '  }',
            '}'
        ].join("\n");

        /* 実行に失敗しても読み込んだアクションは必ず外す / Always unload the action, even if it fails */
        return runTemporaryAction(actionCode, "AreaType", "AutoSize");
    }

    /**
     * エリア内文字に自動サイズ調整を適用する（拡張のみ・OFFには戻さない）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {boolean} 実行できたら true
     */
    function applyAutoSize(doc, textFrame) {
        doc.selection = [textFrame];
        return setAutoSizeByAction(AUTO_SIZE_ON);
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択からエリア内文字を集め、自動サイズ調整を適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.selectObject));
            return;
        }

        /* 処理中に選択が変わるため、先に対象を確定させる / Collect the targets before the selection changes */
        var areaFrames = getSelectedAreaText(doc);
        if (areaFrames.length === 0) {
            alert(getLabel(LABELS.alert.noAreaText));
            return;
        }

        for (var i = 0; i < areaFrames.length; i++) {
            /* 失敗したら残りは処理せずに知らせる / Stop and tell the user on failure */
            if (!applyAutoSize(doc, areaFrames[i])) {
                alert(getLabel(LABELS.alert.actionFailed));
                return;
            }
        }
    }

    main();

})();
