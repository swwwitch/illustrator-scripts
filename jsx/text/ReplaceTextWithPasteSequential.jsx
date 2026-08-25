#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードの複数行テキストから1行目を取り出して、選択中のテキストフレームに適用します（空行は読み飛ばします）。適用した行はクリップボードから取り除かれるため繰り返し実行すると次の行へ順に進み、テキストフレームを選択していないときはウィンドウ中央に貼り付けます。

詳細は README を参照してください。

### Overview

Takes the first line of the multi-line text on the clipboard, skipping blank lines, and applies it to the selected text frames. The applied line is removed from the clipboard, so running the script repeatedly walks through the lines one by one, pasting at the center of the window when no text frame is selected.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceTextWithPasteSequential"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPasteSequential.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPasteSequential.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf4b285b87940"; /* 紹介記事 / article URL */

/**
 * @discussion 原案 / Original idea by Gorolib Design
 * @discussion ReplaceTextWithPaste.jsx から派生 / Derived from ReplaceTextWithPaste.jsx
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
            lastLineApplied: {
                ja: "最後の行を適用しました。\nクリップボードの内容はそのまま残っています。",
                en: "Applied the last line.\nThe clipboard has been left as it is."
            },
            clipboardError: {
                ja: "クリップボードからの取得に失敗しました：\n",
                en: "Failed to get text from clipboard:\n"
            },
            clipboardWriteError: {
                ja: "クリップボードの更新に失敗しました：\n",
                en: "Failed to update the clipboard:\n"
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
     * テキスト編集中の選択（TextRange）から、親のテキストフレームを取り出す。
     * オブジェクト参照は編集モードの解除やペーストで無効になり得るため、
     * フレームだけを取り出して以降の対象にする。
     * @param {Array<Object>} capturedItems - 退避した選択
     * @returns {TextFrame|null} 親のテキストフレーム。編集中でなければ null
     */
    function captureEditingFrame(capturedItems) {
        if (!capturedItems || capturedItems.length !== 1) return null;

        var textRange = capturedItems[0];
        if (!textRange || textRange.typename !== "TextRange") return null;

        try {
            if (textRange.parent && textRange.parent.typename === "TextFrame") return textRange.parent;
            if (textRange.story && textRange.story.textFrames.length > 0) return textRange.story.textFrames[0];
        } catch (e) {
            // 取り出せない場合は編集中として扱わない / Treat it as a normal selection when we cannot tell
        }
        return null;
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
    // クリップボードの読み書き / Clipboard access
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
     * 空白だけの行かどうかを判定する（半角スペース、タブ、全角スペースを空白として扱う）
     * @param {string} lineText - 判定する行
     * @returns {boolean} 空、または空白だけなら true
     */
    function isBlankLine(lineText) {
        return /^[\s　]*$/.test(lineText);
    }

    /**
     * 空行（空白だけの行を含む）を取り除く。
     * 残った行のあいだの改行は、元の文字（段落改行の CR、強制改行の LF）をそのまま使う。
     * @param {string} textContent - 対象の文字列
     * @returns {string} 空行を取り除いた文字列
     */
    function removeEmptyLines(textContent) {
        if (!textContent) return "";

        var keptText = "";
        var lineBuffer = "";
        var breakBeforeLine = "";

        /* 末尾は改行が無くても1行として確定させるため、長さの位置まで回す / Close the last line even without a trailing break */
        for (var i = 0; i <= textContent.length; i++) {
            var isEndOfText = (i === textContent.length);
            var currentChar = isEndOfText ? "" : textContent.charAt(i);

            if (!isEndOfText && currentChar !== "\r" && currentChar !== "\n") {
                lineBuffer += currentChar;
                continue;
            }

            /* 空行は改行ごと捨てるので、残す行の直前の改行だけが書き出される / Dropping a blank line drops its break too */
            if (!isBlankLine(lineBuffer)) {
                keptText += (keptText === "" ? "" : breakBeforeLine) + lineBuffer;
            }
            lineBuffer = "";

            if (isEndOfText) break;

            breakBeforeLine = currentChar;
            /* CR+LF は1つの改行として扱う / Treat CR+LF as a single break */
            if (currentChar === "\r" && textContent.charAt(i + 1) === "\n") {
                breakBeforeLine = "\r\n";
                i++;
            }
        }
        return keptText;
    }

    /**
     * 一度ペーストして、貼り付けられたテキストフレームから文字列を読み取る。
     * 読み取り後は貼り付けたオブジェクトを削除し、元の選択へ戻す。
     * 貼り付け前に選択を解除するのは、ペーストが実行されなかったときに
     * 元の選択を「貼り付いたもの」と誤認して削除しないため。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} originalSelection - 復元する元の選択
     * @returns {string|null} クリップボードの文字列。テキストが無い、または失敗した場合は null
     */
    function readClipboardText(doc, originalSelection) {
        var clipboardText = null;
        var pastedItems = null;
        var pasteError = null;

        try {
            /* 貼り付いたものだけを確実に拾うため、先に選択を空にする / Clear the selection first so only the pasted items are captured */
            setSelection(doc, null);
            pastedItems = pasteClipboardItems(doc);

            var pastedTextFrame = findFirstTextFrame(pastedItems);
            if (pastedTextFrame) {
                /* ここで空行を落としておけば、1行目の取り出しも書き戻しも空行を意識しなくてよい / Strip blank lines once, so neither the split nor the write-back has to care */
                var pastedText = removeEmptyLines(pastedTextFrame.contents);
                /* 中身が空なら「残りなし」として扱う / Nothing usable means there is nothing left to apply */
                if (pastedText.length > 0) {
                    clipboardText = pastedText;
                }
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
        } else if (clipboardText === null) {
            alert(getLabel("alert.noTextInClipboard"));
        }
        return clipboardText;
    }

    /**
     * 一時テキストフレーム経由でクリップボードを書き換える。
     * Illustrator には文字列を直接クリップボードへ送る API が無いため、
     * 内容を持つフレームを作ってコピーし、すぐに削除する。
     * 追加直後のフレームは再描画しないとコピー対象として扱われないことがあり、
     * また app.copy() は黙って無視される場合があるためメニューコマンドを使う。
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - クリップボードに残す文字列
     * @returns {void}
     */
    function writeTextToClipboard(doc, textContent) {
        var tempFrame = null;
        var writeError = null;

        try {
            tempFrame = doc.activeLayer.textFrames.add();
            tempFrame.contents = textContent;

            /* 追加したフレームを画面に反映してから選択する / Flush the new frame to the canvas before selecting it */
            app.redraw();
            app.executeMenuCommand("deselectall");
            tempFrame.selected = true;
            app.redraw();

            app.executeMenuCommand("copy");
            /* コピーが確定してから元のフレームを消す / Let the copy settle before deleting the source */
            app.redraw();
        } catch (e) {
            writeError = String(e);
        }

        if (tempFrame) removeItems([tempFrame]);
        if (writeError) {
            alert(getLabel("alert.clipboardWriteError") + writeError);
        }
    }

    // =========================================
    // ウィンドウ中央への配置 / Placement at the center of the window
    // =========================================

    /**
     * 表示中のウィンドウの中心座標を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{x: number, y: number}} ウィンドウ中心の座標
     */
    function getActiveViewCenter(doc) {
        var viewCenter = doc.activeView.centerPoint;
        return { x: viewCenter[0], y: viewCenter[1] };
    }

    /**
     * オブジェクト群を囲む矩形の中心座標を返す
     * @param {Array<Object>} targetItems - 対象のページアイテム
     * @returns {{x: number, y: number}|null} 中心座標。座標を取れるものが無ければ null
     */
    function getItemsCenter(targetItems) {
        var left = null, top = null, right = null, bottom = null;

        for (var i = 0; i < targetItems.length; i++) {
            var bounds = null;
            try {
                bounds = targetItems[i].geometricBounds;
            } catch (e) {
                /* TextRange など座標を持たないものは対象外にする / Skip what has no geometry, such as a TextRange */
                continue;
            }
            if (!bounds || bounds.length < 4) continue;

            if (left === null || bounds[0] < left) left = bounds[0];
            if (top === null || bounds[1] > top) top = bounds[1];
            if (right === null || bounds[2] > right) right = bounds[2];
            if (bottom === null || bounds[3] < bottom) bottom = bounds[3];
        }

        if (left === null) return null;
        return { x: (left + right) / 2, y: (top + bottom) / 2 };
    }

    /**
     * オブジェクト群をまとめてウィンドウの中央へ移動する。
     * 複数まとめて貼り付いた場合も並びを崩さないよう、全体を同じ量だけ動かす。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} targetItems - 移動するページアイテム
     * @returns {void}
     */
    function moveItemsToViewCenter(doc, targetItems) {
        var itemsCenter = getItemsCenter(targetItems);
        if (!itemsCenter) return;

        var viewCenter = getActiveViewCenter(doc);
        var offsetX = viewCenter.x - itemsCenter.x;
        var offsetY = viewCenter.y - itemsCenter.y;

        for (var i = 0; i < targetItems.length; i++) {
            try {
                targetItems[i].translate(offsetX, offsetY);
            } catch (e) {
                // 動かせないものはその位置に残す / Leave behind whatever cannot be moved
            }
        }
    }

    /**
     * 貼り付いたテキストフレームの内容を1行目だけに切り詰め、2行目以降を返す。
     * @param {TextFrame} pastedTextFrame - 貼り付いたテキストフレーム
     * @returns {string|null} 2行目以降の文字列。1行しかなければ空文字、使えるテキストが無ければ null
     */
    function trimPastedTextToFirstLine(pastedTextFrame) {
        /* ここで空行を落としておけば、1行目の取り出しも書き戻しも空行を意識しなくてよい / Strip blank lines once, so neither the split nor the write-back has to care */
        var pastedText = removeEmptyLines(pastedTextFrame.contents);
        if (pastedText.length === 0) return null;

        var splitResult = splitFirstLine(pastedText);
        pastedTextFrame.contents = splitResult.firstLine;
        /* 切り詰めた分を座標へ反映させてから中央を測る / Flush the trim before the bounds are measured */
        app.redraw();
        return splitResult.remainder;
    }

    /**
     * クリップボードの1行目だけをウィンドウの中央へ貼り付ける。
     * 適用先のテキストフレームが無いときの動作で、貼り付いたフレームを1行目だけに切り詰め、
     * 2行目以降はクリップボードへ書き戻して次の実行に引き継ぐ。
     * ペースト位置もウィンドウの中央だが、1行目に切り詰めるとフレームの大きさが変わるため、
     * 切り詰めたあとに測り直して中央へ置き直す。
     * テキストを含まない内容は切り詰めも書き戻しも行わず、そのまま中央へ置く。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function pasteFirstLineAtViewCenter(doc) {
        var pastedItems = null;
        var pasteError = null;

        try {
            pastedItems = pasteClipboardItems(doc);
        } catch (e) {
            pasteError = String(e);
        }

        if (pasteError) {
            alert(getLabel("alert.clipboardError") + pasteError);
            return;
        }
        if (!pastedItems || pastedItems.length === 0) {
            alert(getLabel("alert.emptyClipboard"));
            return;
        }

        /* テキスト以外は取り出す行が無いので、貼り付いたまま中央へ置く / Non-text has no line to take, so leave the paste as it is */
        var pastedTextFrame = findFirstTextFrame(pastedItems);
        var remainder = pastedTextFrame ? trimPastedTextToFirstLine(pastedTextFrame) : "";

        if (remainder === null) {
            /* 空行だけの内容は、貼り付けた分を片付けてから知らせる / Clean up before reporting a paste of blank lines only */
            removeItems(pastedItems);
            setSelection(doc, null);
            alert(getLabel("alert.noTextInClipboard"));
            return;
        }

        moveItemsToViewCenter(doc, pastedItems);

        /* 貼り付けた行を取り除いた残りをクリップボードへ戻し、次の実行で続きから進めるようにする / Put the remaining lines back so the next run continues where this one stopped */
        var hasRemainder = (remainder.length > 0);
        if (hasRemainder) {
            writeTextToClipboard(doc, remainder);
        }

        /* 貼り付けた直後に手を加えられるよう選択したままにする / Keep the paste selected so it can be edited right away */
        setSelection(doc, pastedItems);
        app.redraw();

        if (pastedTextFrame && !hasRemainder) {
            alert(getLabel("alert.lastLineApplied"));
        }
    }

    // =========================================
    // テキストの適用 / Text application
    // =========================================

    /**
     * 文字列を最初の改行で1行目と残りに分ける。
     * Illustrator のテキストは段落改行が CR、強制改行が LF になるため、どちらも区切りとして扱う。
     * @param {string} textContent - 分割する文字列
     * @returns {{firstLine: string, remainder: string}} 1行目と、改行を取り除いた残り
     */
    function splitFirstLine(textContent) {
        var lineBreak = /\r\n|\r|\n/.exec(textContent);
        if (!lineBreak) {
            return { firstLine: textContent, remainder: "" };
        }
        return {
            firstLine: textContent.substring(0, lineBreak.index),
            remainder: textContent.substring(lineBreak.index + lineBreak[0].length)
        };
    }

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
     * クリップボードの1行目を選択中のテキストフレームへ適用し、残りをクリップボードへ書き戻す。
     * 適用先が無い場合は、1行目をウィンドウ中央へ貼り付ける。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        /* 文字を選択して編集中なら、編集を抜けて親テキストフレームを対象にする / Leave text editing and target the parent frame */
        var editingFrame = captureEditingFrame(originalSelection);
        if (editingFrame) {
            leaveTextEditing(doc);
            originalSelection = [editingFrame];
            setSelection(doc, originalSelection);
        }

        /* グループやクリップグループの中は、ペーストを挟んで参照が古くなる前にたどっておく / Walk into groups before the paste cycle can stale the references */
        var targetFrames = collectTextFramesFrom(originalSelection);
        /* 適用先が無いときは、1行目をウィンドウ中央へ貼り付けて残りをクリップボードへ戻す / Without a target, paste the first line at the center of the window and keep the rest */
        if (targetFrames.length === 0) {
            pasteFirstLineAtViewCenter(doc);
            return;
        }

        var clipboardText = readClipboardText(doc, originalSelection);
        if (clipboardText === null) return;

        var splitResult = splitFirstLine(clipboardText);

        var errorMessages = [];
        applyTextToFrames(targetFrames, splitResult.firstLine, errorMessages);
        /* 選択数だけダイアログが出ないよう、まとめて1回だけ知らせる / Report every failure in a single alert */
        if (errorMessages.length > 0) {
            alert(getLabel("alert.replaceError") + errorMessages.join("\n"));
        }

        /* 適用した行を取り除いた残りをクリップボードへ戻し、次の実行で続きから進めるようにする / Put the remaining lines back so the next run continues where this one stopped */
        var hasRemainder = (splitResult.remainder.length > 0);
        if (hasRemainder) {
            writeTextToClipboard(doc, splitResult.remainder);
        }

        /* 置換後の表示を確実に更新するため、選択を解除してから戻す / Clear and reset the selection so the redraw reflects the change */
        setSelection(doc, null);
        app.redraw();
        setSelection(doc, originalSelection);

        if (!hasRemainder) {
            alert(getLabel("alert.lastLineApplied"));
        }
    }

    main();

})();
