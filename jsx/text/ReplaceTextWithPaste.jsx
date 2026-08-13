#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードのテキストで、選択中のテキストフレームの内容をまとめて置き換えます。選択がない場合は、貼り付いた位置に既定の書式でテキストフレームを新規作成します。

詳細は README を参照してください。

### Overview

Replaces the contents of the selected text frames with the text on the clipboard. With nothing selected, it creates a new text frame with the default formatting where the paste lands.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceTextWithPaste";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-14";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPaste.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPaste.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf14ce08eb618"; /* 紹介記事 / article URL */

/**
 * @discussion 原案 / Original idea by Gorolib Design
 */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールから表示言語を決める / Pick the UI language from the locale */
    var currentLanguage = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextInClipboard: {
                ja: "クリップボードにテキストが見つかりませんでした。",
                en: "No text was found on the clipboard."
            },
            clipboardError: {
                ja: "クリップボードからの取得に失敗しました：\n",
                en: "Failed to get text from clipboard:\n"
            },
            textCreateError: {
                ja: "新規テキスト作成時にエラーが発生しました：\n",
                en: "An error occurred while creating new text:\n"
            },
            replaceError: {
                ja: "テキスト置換中にエラーが発生しました：\n",
                en: "An error occurred while replacing text:\n"
            }
        }
    };

    /**
     * ラベルをドット区切りのキーで引く
     * @param {string} labelKey - "alert.noDocument" のようなドット区切りキー
     * @returns {string} 現在の表示言語のラベル。見つからない場合はキーをそのまま返す
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) return labelKey;
            labelNode = labelNode[keyParts[i]];
        }
        return (labelNode && labelNode[currentLanguage]) ? labelNode[currentLanguage] : labelKey;
    }

    // =========================================
    // 選択の操作 / Selection helpers
    // =========================================

    /**
     * 選択オブジェクトを配列に写し取る（selection は操作で変化するため）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<Object>} 選択オブジェクトの配列。選択がなければ空配列
     */
    function captureSelection(doc) {
        var capturedItems = [];
        var currentSelection = doc.selection;
        if (!currentSelection || !currentSelection.length) return capturedItems;
        for (var i = 0; i < currentSelection.length; i++) {
            capturedItems.push(currentSelection[i]);
        }
        return capturedItems;
    }

    /**
     * 選択状態を差し替える（失敗しても処理を止めない）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} itemsToSelect - 選択するオブジェクトの配列。空または null で選択解除
     * @returns {void}
     */
    function setSelection(doc, itemsToSelect) {
        try {
            doc.selection = (itemsToSelect && itemsToSelect.length > 0) ? itemsToSelect : null;
        } catch (e) {
            // Illustrator が選択を拒む場合は現在の選択のままにする / Keep whatever stays selected
        }
    }

    /**
     * 指定のオブジェクトをまとめて削除する
     * @param {Array<Object>} itemsToRemove - 削除対象の配列
     * @returns {void}
     */
    function removeItems(itemsToRemove) {
        if (!itemsToRemove || !itemsToRemove.length) return;
        for (var i = itemsToRemove.length - 1; i >= 0; i--) {
            try {
                /* TextRange の remove() は文字そのものを消すため対象外にする / TextRange.remove() would delete characters, so skip it */
                if (itemsToRemove[i].typename === "TextRange") continue;
                itemsToRemove[i].remove();
            } catch (e) {
                // 既に消えているものは無視する / Ignore items that are already gone
            }
        }
    }

    /**
     * 配列から最初のテキストフレームを探す
     * @param {Array<Object>} searchItems - 探索対象の配列
     * @returns {TextFrame|null} 見つかったテキストフレーム。なければ null
     */
    function findFirstTextFrame(searchItems) {
        if (!searchItems) return null;
        for (var i = 0; i < searchItems.length; i++) {
            if (searchItems[i].typename === "TextFrame") return searchItems[i];
        }
        return null;
    }

    // =========================================
    // クリップボードの取得 / Clipboard access
    // =========================================

    /**
     * 一度ペーストして、貼り付けられたテキストフレームから内容と座標を読み取る。
     * 読み取り後は貼り付けたオブジェクトを削除し、元の選択へ戻す。
     * 貼り付け前に選択を解除するのは、ペーストが実行されなかったときに
     * 元の選択を「貼り付いたもの」と誤認して削除しないため。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} originalSelection - 復元する元の選択
     * @returns {{bounds: Array<number>, contents: string}|null} 読み取り結果。テキストが無い、または失敗した場合は null
     */
    function readClipboardTextFrame(doc, originalSelection) {
        var clipboardInfo = null;
        var pastedItems = null;
        var pasteError = null;

        try {
            /* 貼り付いたものだけを確実に拾うため、先に選択を空にする / Clear the selection first so only the pasted items are captured */
            setSelection(doc, null);
            app.paste();
            pastedItems = captureSelection(doc);

            var pastedTextFrame = findFirstTextFrame(pastedItems);
            if (pastedTextFrame) {
                clipboardInfo = {
                    bounds: pastedTextFrame.geometricBounds,
                    contents: pastedTextFrame.contents
                };
            }
        } catch (e) {
            pasteError = String(e);
        }

        /* 成否にかかわらず、貼り付けた分を消して元の選択へ戻す / Clean up and restore regardless of the outcome */
        removeItems(pastedItems);
        setSelection(doc, originalSelection);

        /* 画面を元に戻してから知らせる / Report only after the canvas is back to its original state */
        if (pasteError) {
            alert(getLabel("alert.clipboardError") + pasteError);
        } else if (!clipboardInfo) {
            alert(getLabel("alert.noTextInClipboard"));
        }
        return clipboardInfo;
    }

    // =========================================
    // テキストの適用 / Text application
    // =========================================

    /**
     * 同じ内容のエラーを重複させずに追加する
     * @param {Array<string>} errorMessages - 収集先
     * @param {string} errorMessage - 追加するエラーメッセージ
     * @returns {void}
     */
    function addUniqueError(errorMessages, errorMessage) {
        for (var i = 0; i < errorMessages.length; i++) {
            if (errorMessages[i] === errorMessage) return;
        }
        errorMessages.push(errorMessage);
    }

    /**
     * クリップボードのテキストを、貼り付いた位置に新規テキストフレームとして作成する
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {Array<number>} pastedBounds - 貼り付いたテキストフレームの geometricBounds
     * @param {string} textContent - 作成するテキスト
     * @returns {void}
     */
    function createNewTextFrame(targetLayer, pastedBounds, textContent) {
        try {
            var newFrame = targetLayer.textFrames.add();
            newFrame.contents = textContent;

            var newFrameBounds = newFrame.geometricBounds;
            newFrame.translate(pastedBounds[0] - newFrameBounds[0], pastedBounds[1] - newFrameBounds[1]);
        } catch (e) {
            alert(getLabel("alert.textCreateError") + e);
        }
    }

    /**
     * テキストフレームなら内容を置き換え、グループなら中を再帰的にたどる。
     * TextRange が選択されている場合は親のテキストフレームに読み替える。
     * @param {Object} targetItem - 対象のページアイテムまたは TextRange
     * @param {string} textContent - 適用するテキスト
     * @param {Array<string>} errorMessages - 発生したエラーの収集先（呼び出し元でまとめて通知する）
     * @returns {void}
     */
    function applyTextToItem(targetItem, textContent, errorMessages) {
        try {
            if (targetItem.typename === "TextRange" && targetItem.parent.typename === "TextFrame") {
                targetItem = targetItem.parent;
            }

            if (targetItem.typename === "TextFrame") {
                targetItem.contents = textContent;
                return;
            }

            if (targetItem.typename === "GroupItem") {
                var childItems = targetItem.pageItems;
                for (var i = 0; i < childItems.length; i++) {
                    applyTextToItem(childItems[i], textContent, errorMessages);
                }
            }
        } catch (e) {
            addUniqueError(errorMessages, String(e));
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の退避、クリップボードの取得、テキストの適用までを通して行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        var clipboardInfo = readClipboardTextFrame(doc, originalSelection);
        if (!clipboardInfo) return;

        if (originalSelection.length === 0) {
            createNewTextFrame(doc.activeLayer, clipboardInfo.bounds, clipboardInfo.contents);
        } else {
            var errorMessages = [];
            for (var i = 0; i < originalSelection.length; i++) {
                applyTextToItem(originalSelection[i], clipboardInfo.contents, errorMessages);
            }
            /* 選択数だけダイアログが出ないよう、まとめて1回だけ知らせる / Report every failure in a single alert */
            if (errorMessages.length > 0) {
                alert(getLabel("alert.replaceError") + errorMessages.join("\n"));
            }
        }

        /* 置換後の表示を確実に更新するため、選択を解除してから戻す / Clear and reset the selection so the redraw reflects the change */
        setSelection(doc, null);
        app.redraw();
        setSelection(doc, originalSelection);
    }

    main();

})();
