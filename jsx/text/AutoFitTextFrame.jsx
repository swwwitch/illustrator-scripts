#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字に自動サイズ調整を適用します（拡張のみ）。
FitAreaText の［自動サイズ調整］だけを、ダイアログなしで実行する版です。

### Overview

Applies Auto Size to the selected area type (expand only).
A dialog-free version of the Auto Size option in FitAreaText.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoFitTextFrame";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

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
            }
        }
    };

    /**
     * 現在の言語のラベルを取得する
     * @param {Object} labelSet - ja/en を持つラベル
     * @returns {string} 表示用の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

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

    // =========================================
    // 自動サイズ調整 / Auto Size
    // =========================================

    /**
     * 自動サイズ調整をアクション経由で切り替える
     * @param {number} autoSizeValue - 1（ON）または 2（OFF）
     * @returns {void}
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

        var actionFile = new File('~/ScriptAction.aia');
        actionFile.open('w');
        actionFile.write(actionCode);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();

        /* 実行に失敗しても読み込んだアクションは必ず外す / Always unload the action, even if it fails */
        try {
            app.doScript("AutoSize", "AreaType", false);
        } finally {
            app.unloadAction("AreaType", "");
        }
    }

    /**
     * エリア内文字に自動サイズ調整を適用する（拡張のみ・OFFには戻さない）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function applyAutoSize(doc, textFrame) {
        doc.selection = [textFrame];
        setAutoSizeByAction(AUTO_SIZE_ON);
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
            applyAutoSize(doc, areaFrames[i]);
        }
    }

    main();

})();
