#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードのテキストで、選択中のテキストフレームの内容をまとめて置き換えます。選択がない場合は、貼り付いた位置（画面の中央）に既定の書式でテキストフレームを新規作成します。

詳細は README を参照してください。

### Overview

Replaces the contents of the selected text frames with the text on the clipboard. With nothing selected, it creates a new text frame with the default formatting where the paste lands, which is the center of the view.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceTextWithPaste";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

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
            emptyClipboard: {
                ja: "クリップボードが空か、Illustrator に貼り付けられない内容です。",
                en: "The clipboard is empty, or Illustrator cannot paste its contents."
            },
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
     * 選択オブジェクトを配列に写し取る（selection は操作で変化するため）。
     * テキスト編集中は selection が配列ではなく TextRange そのものになることがあり、
     * その length は選択した文字数を指す。添字で取り出すと undefined が並び、
     * あとで選択に戻すときに Illustrator が落ちるため、形を判定してから写す。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<Object>} 選択オブジェクトの配列。選択がなければ空配列
     */
    function captureSelection(doc) {
        var capturedItems = [];
        var currentSelection = doc.selection;
        if (!currentSelection) return capturedItems;

        if (!(currentSelection instanceof Array)) {
            if (currentSelection.typename === "TextRange") capturedItems.push(currentSelection);
            return capturedItems;
        }

        for (var i = 0; i < currentSelection.length; i++) {
            if (currentSelection[i]) capturedItems.push(currentSelection[i]);
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
     * テキスト編集中の選択（TextRange）を、あとから使える数値として控える。
     * オブジェクト参照は編集モードの解除やペーストで無効になり得るため、
     * 親テキストフレームと文字位置だけを持たせる。
     * @param {Array<Object>} capturedItems - 退避した選択
     * @returns {{frame: TextFrame, start: number, end: number}|null} 文字範囲の情報。編集中でなければ null
     */
    function captureEditingRange(capturedItems) {
        if (!capturedItems || capturedItems.length !== 1) return null;

        var textRange = capturedItems[0];
        if (!textRange || textRange.typename !== "TextRange") return null;

        try {
            var parentFrame = null;
            if (textRange.parent && textRange.parent.typename === "TextFrame") {
                parentFrame = textRange.parent;
            } else if (textRange.story && textRange.story.textFrames.length > 0) {
                parentFrame = textRange.story.textFrames[0];
            }
            if (!parentFrame) return null;

            return { frame: parentFrame, start: textRange.start, end: textRange.end };
        } catch (e) {
            return null;
        }
    }

    /**
     * テキスト編集モードを抜ける。
     * 編集中のままペーストするとクリップボードの内容が文字として流し込まれ、
     * さらに再描画を挟むと Illustrator が不安定になるため、読み取りの前に必ず抜ける。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function leaveTextEditing(doc) {
        try {
            /* 選択ツールへ切り替えると編集が確定して抜けられる / Switching tools commits the edit and leaves it */
            app.selectTool("Adobe Select Tool");
        } catch (e) {
            // 切り替えられない場合は次の選択解除に任せる / Leave it to the deselect below
        }
        setSelection(doc, null);
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
     * テキストフレームを再帰的に集める。
     * クリップグループも typename は GroupItem なので、同じ経路で中までたどれる。
     * 子は生のコレクションのまま持たず配列へ写し取る。内容を書き換えると
     * コレクションの中身が変わり、たどっている途中で取りこぼすため。
     * TextRange が渡された場合は親のテキストフレームに読み替える。
     * @param {Object} searchItem - 探索対象のページアイテムまたは TextRange
     * @param {Array<TextFrame>} collectedFrames - 収集先の配列
     * @returns {void}
     */
    function collectTextFrames(searchItem, collectedFrames) {
        if (!searchItem) return;

        try {
            if (searchItem.typename === "TextRange" && searchItem.parent && searchItem.parent.typename === "TextFrame") {
                searchItem = searchItem.parent;
            }

            if (searchItem.typename === "TextFrame") {
                collectedFrames.push(searchItem);
                return;
            }

            if (searchItem.typename !== "GroupItem") return;

            var childItems = [];
            for (var i = 0; i < searchItem.pageItems.length; i++) {
                childItems.push(searchItem.pageItems[i]);
            }
            for (var j = 0; j < childItems.length; j++) {
                collectTextFrames(childItems[j], collectedFrames);
            }
        } catch (e) {
            // シンボルやエンベロープなど、中をたどれないものは対象外にする / Skip what we cannot walk into, such as symbols and envelopes
        }
    }

    /**
     * 配列やコレクションから、テキストフレームをまとめて集める
     * @param {Array<Object>} searchItems - 探索対象の配列またはコレクション
     * @returns {Array<TextFrame>} 見つかったテキストフレームの配列。順序は探索順
     */
    function collectTextFramesFrom(searchItems) {
        var collectedFrames = [];
        if (!searchItems) return collectedFrames;
        for (var i = 0; i < searchItems.length; i++) {
            collectTextFrames(searchItems[i], collectedFrames);
        }
        return collectedFrames;
    }

    /**
     * 配列から最初のテキストフレームを探す。
     * 他アプリからのペーストはグループやクリップグループにまとめられることがあるため、中も再帰的にたどる。
     * @param {Array<Object>} searchItems - 探索対象の配列またはコレクション
     * @returns {TextFrame|null} 見つかったテキストフレーム。なければ null
     */
    function findFirstTextFrame(searchItems) {
        var foundFrames = collectTextFramesFrom(searchItems);
        return (foundFrames.length > 0) ? foundFrames[0] : null;
    }

    // =========================================
    // クリップボードの取得 / Clipboard access
    // =========================================

    /**
     * クリップボードの内容をドキュメントへ貼り付け、貼り付いたオブジェクトを返す。
     * Illustrator は自分がコピーした内容を内部に保持していて、他アプリがクリップボードを
     * 書き換えたあとの1回目のペーストでは古い内容が貼り付く。その1回目が内部の更新を促すため、
     * 1回目は捨てて2回目の結果を使う。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<Object>} 貼り付いたオブジェクトの配列。貼り付かなかった場合は空配列
     */
    function pasteClipboardItems(doc) {
        try {
            /* 1回目は内部クリップボードを最新にするためだけのペースト / The first paste only refreshes Illustrator's cached clipboard */
            setSelection(doc, null);
            app.paste();
            app.redraw();
        } catch (e) {
            // 更新目的なので、失敗しても2回目の結果で判断する / Judge by the second paste even if this one fails
        }
        removeItems(captureSelection(doc));

        setSelection(doc, null);
        app.paste();
        /* 貼り付け直後は selection に反映されないことがあるため、描画を確定させてから読む / Flush the paste before reading the selection */
        app.redraw();
        return captureSelection(doc);
    }

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
            pastedItems = pasteClipboardItems(doc);

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
        } else if (!pastedItems || pastedItems.length === 0) {
            /* ペースト自体が起きなかった場合と、貼り付いたがテキストが無い場合を区別する / Tell an unusable clipboard apart from a paste without text */
            alert(getLabel("alert.emptyClipboard"));
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
     * 控えた文字範囲だけをクリップボードのテキストで置き換える。
     * カーソルがあるだけで文字が選ばれていない場合は、そのテキストフレーム全体を対象にする。
     * @param {{frame: TextFrame, start: number, end: number}} rangeInfo - 置き換える範囲
     * @param {string} textContent - 適用するテキスト
     * @param {Array<string>} errorMessages - 発生したエラーの収集先（呼び出し元でまとめて通知する）
     * @returns {void}
     */
    function replaceEditingRange(rangeInfo, textContent, errorMessages) {
        if (rangeInfo.start === rangeInfo.end) {
            applyTextToFrames([rangeInfo.frame], textContent, errorMessages);
            return;
        }

        try {
            var targetRange = rangeInfo.frame.textRange;
            targetRange.start = rangeInfo.start;
            targetRange.end = rangeInfo.end;
            targetRange.contents = textContent;
        } catch (e) {
            addUniqueError(errorMessages, String(e));
        }
    }

    /**
     * 集めたテキストフレームの内容をまとめて置き換える
     * @param {Array<TextFrame>} targetFrames - 対象のテキストフレーム
     * @param {string} textContent - 適用するテキスト
     * @param {Array<string>} errorMessages - 発生したエラーの収集先（呼び出し元でまとめて通知する）
     * @returns {void}
     */
    function applyTextToFrames(targetFrames, textContent, errorMessages) {
        for (var i = 0; i < targetFrames.length; i++) {
            try {
                targetFrames[i].contents = textContent;
            } catch (e) {
                addUniqueError(errorMessages, String(e));
            }
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

        /* 文字を選択して編集中なら、その範囲だけを置き換える / When characters are selected, replace just that range */
        var editingRange = captureEditingRange(originalSelection);
        if (editingRange) {
            leaveTextEditing(doc);

            var editingClipboard = readClipboardTextFrame(doc, []);
            if (!editingClipboard) return;

            var editingErrors = [];
            replaceEditingRange(editingRange, editingClipboard.contents, editingErrors);
            if (editingErrors.length > 0) {
                alert(getLabel("alert.replaceError") + editingErrors.join("\n"));
            }

            /* 編集モードは抜けているので、対象のテキストフレームを選択して終える / Editing is over, so leave the frame itself selected */
            setSelection(doc, [editingRange.frame]);
            app.redraw();
            return;
        }

        /* グループやクリップグループの中は、ペーストを挟んで参照が古くなる前にたどっておく / Walk into groups before the paste cycle can stale the references */
        var targetFrames = collectTextFramesFrom(originalSelection);

        var clipboardInfo = readClipboardTextFrame(doc, originalSelection);
        if (!clipboardInfo) return;

        if (originalSelection.length === 0) {
            createNewTextFrame(doc.activeLayer, clipboardInfo.bounds, clipboardInfo.contents);
        } else {
            var errorMessages = [];
            applyTextToFrames(targetFrames, clipboardInfo.contents, errorMessages);
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
