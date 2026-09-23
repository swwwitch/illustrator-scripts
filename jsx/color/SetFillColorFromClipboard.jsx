#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードにある「204,225,179」のようなRGB文字列を読み取り、塗り色に設定します。
ReplaceTextWithPaste.jsx と同じ2回ペースト方式で、1回目は削除、2回目で値を読み取ってからカットします。

### 注意

カットによって、クリップボードの内容はIllustratorのテキストオブジェクトに置き換わります。

### Overview

Reads an RGB string such as "204,225,179" from the clipboard and applies it as the fill color.
As in ReplaceTextWithPaste.jsx it pastes twice: the first paste is discarded, the second is read and then cut.

### Notes

The cut replaces the clipboard contents with an Illustrator text object.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetFillColorFromClipboard";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            textEditing: {
                ja: "文字の編集を終了してから実行してください。",
                en: "Please exit text editing before running this script."
            },
            notSingleText: {
                ja: "RGB値が書かれたテキストを1つだけコピーしてください。",
                en: "Please copy exactly one text object with an RGB value in it."
            },
            invalidFormat: {
                ja: "「204,225,179」の形式でRGB値をコピーしてください。",
                en: "Please copy an RGB value in the “204,225,179” format."
            },
            outOfRange: { ja: "RGB値は0〜255で指定してください。", en: "RGB values must be between 0 and 255." },
            failed: { ja: "塗り色を設定できませんでした。", en: "Could not set the fill color." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var initialSelection = doc.selection;

    /* TextRangeは編集終了後に無効になるため保持しない / A TextRange becomes invalid once editing ends, so it is not kept */
    if (initialSelection && initialSelection.typename === "TextRange") {
        alert(getLabel("alert.textEditing"));
        return;
    }

    var originalSelection = toItemArray(initialSelection);
    var pastedItems = [];
    var rgbValues = null;
    var failureError = null;

    try {
        rgbValues = readRgbFromClipboard();
    } catch (e) {
        failureError = e;
    } finally {
        /* 不正な値やカット失敗の場合も、一時オブジェクトを残さない / Never leave temporary objects behind */
        removePastedItems();
        doc.selection = originalSelection.length ? originalSelection : null;
    }

    if (failureError) {
        alert(getLabel("alert.failed") + "\n" + (failureError.message || failureError));
        return;
    }

    var fillColor = new RGBColor();
    fillColor.red = rgbValues[0];
    fillColor.green = rgbValues[1];
    fillColor.blue = rgbValues[2];
    doc.defaultFillColor = fillColor;
    app.redraw();

    // =========================================
    // 一時オブジェクトの操作 / Temporary items
    // =========================================

    /**
     * 選択をページアイテムの配列に変換する
     * @param {Object} selectionSource - doc.selection などの選択
     * @returns {PageItem[]} 有効なページアイテムの配列
     */
    function toItemArray(selectionSource) {
        var items = [];
        if (!selectionSource) return items;
        for (var i = 0; i < selectionSource.length; i++) {
            if (selectionSource[i]) items.push(selectionSource[i]);
        }
        return items;
    }

    /**
     * クリップボードを貼り付けて、生成されたオブジェクトを pastedItems に取り込む
     * @returns {void}
     */
    function pasteAndCapture() {
        /* 解除しないと選択中のオブジェクトや文字が置き換わる / Clearing the selection first avoids replacing it */
        doc.selection = null;
        app.paste();
        app.redraw();

        var pastedSelection = doc.selection;
        if (pastedSelection && pastedSelection.typename === "TextRange") {
            throw new Error(getLabel("alert.textEditing"));
        }
        pastedItems = toItemArray(pastedSelection);
    }

    /**
     * 取り込んだ一時オブジェクトを削除する
     * @returns {void}
     */
    function removePastedItems() {
        for (var i = pastedItems.length - 1; i >= 0; i--) {
            /* カット済みの参照は無効なので無視する / References already cut are invalid */
            try {
                pastedItems[i].remove();
            } catch (e) {}
        }
        pastedItems = [];
    }

    // =========================================
    // RGB値の読み取り / Reading the RGB value
    // =========================================

    /**
     * ページアイテムに含まれる文字列を再帰的に集める
     * @param {PageItem} item - 対象のページアイテム
     * @param {string[]} collectedTexts - 収集先の配列
     * @returns {void}
     */
    function collectTextContents(item, collectedTexts) {
        if (item.typename === "TextFrame") {
            collectedTexts.push(item.contents);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectTextContents(item.pageItems[i], collectedTexts);
            }
        }
    }

    /**
     * 「204,225,179」形式の文字列をRGB値に変換する
     * @param {string} rgbText - 変換元の文字列
     * @returns {number[]} R・G・Bの3要素
     */
    function parseRgbText(rgbText) {
        var rgbMatch = /^\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*$/.exec(rgbText);
        if (!rgbMatch) {
            throw new Error(getLabel("alert.invalidFormat"));
        }
        var parsedValues = [Number(rgbMatch[1]), Number(rgbMatch[2]), Number(rgbMatch[3])];
        for (var i = 0; i < parsedValues.length; i++) {
            if (parsedValues[i] > 255) throw new Error(getLabel("alert.outOfRange"));
        }
        return parsedValues;
    }

    /**
     * クリップボードを2回貼り付けてRGB値を読み取り、一時オブジェクトをカットする
     * @returns {number[]} R・G・Bの3要素
     */
    function readRgbFromClipboard() {
        app.selectTool("Adobe Select Tool");

        /* 1回目は内部クリップボードの更新用。カットせず削除する / The first paste only refreshes the internal clipboard */
        pasteAndCapture();
        removePastedItems();

        /* 2回目で値を読み取る / The second paste is the one that is read */
        pasteAndCapture();
        var collectedTexts = [];
        for (var i = 0; i < pastedItems.length; i++) {
            collectTextContents(pastedItems[i], collectedTexts);
        }
        if (collectedTexts.length !== 1) {
            throw new Error(getLabel("alert.notSingleText"));
        }

        var parsedValues = parseRgbText(collectedTexts[0]);
        /* 値を読み取った後で、貼り付けたオブジェクトだけをカット / Cut only the pasted objects, once the value is read */
        doc.selection = pastedItems;
        app.cut();
        return parsedValues;
    }
})();
