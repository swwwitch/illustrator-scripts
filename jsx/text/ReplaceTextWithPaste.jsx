#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- クリップボードのテキストで、選択中のテキストフレームの内容をまとめて置き換えます。
- 選択がない場合は、コピー元と同じ位置に新しいテキストフレームを作成します。
- グループを選択した場合は、その中のテキストフレームをすべて対象にします。
- 日本語／英語インターフェースに対応します。

### 主な機能

- 選択したテキストフレームの内容を一括置換
- 選択なしのときは、コピー元と同じ位置に新規テキストフレームを作成
- グループ内のテキストフレームを再帰的に処理
- ポイント文字・エリア内文字・パス上文字に対応
- 処理の前後で選択状態を保持

### 処理の流れ

1. 現在の選択を退避する
2. 一度通常のペーストを実行し、貼り付けられたテキストフレームから内容と座標を取得する
3. 貼り付けたオブジェクトを削除し、退避した選択を復元する
4. 選択があれば各テキストフレームの内容を置き換え、なければ新規テキストフレームを作成する

### 対象

- 対象：テキストフレーム、グループ内のテキストフレーム
- 非対象：画像、図形、ロックされたオブジェクト

### 注意

- Illustrator には取り消しをグループ化する API がないため、この処理の取り消しは複数ステップに分かれます。

*/

/*

### Overview

- Replaces the contents of the selected text frames with the text on the clipboard.
- With nothing selected, creates a new text frame where the text was copied from.
- When a group is selected, every text frame inside it is processed.
- Japanese and English user interface.

### Main Features

- Replaces the contents of every selected text frame at once
- Creates a new text frame at the original position when nothing is selected
- Walks into groups and processes the text frames inside them
- Supports point type, area type, and type on a path
- Keeps the selection intact across the run

### Workflow

1. Save the current selection.
2. Run a normal paste once, then read the contents and bounds from the pasted text frame.
3. Remove the pasted objects and restore the saved selection.
4. Replace the contents of the selected text frames, or create a new one when nothing was selected.

### Scope

- Handled: text frames, and text frames inside groups
- Not handled: images, shapes, locked objects

### Note

- Illustrator has no undo-grouping API, so undoing this run takes multiple steps.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceTextWithPaste";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */
var SCRIPT_CREDIT   = "Original idea by Gorolib Design"; /* 原案 / original idea */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceTextWithPaste.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPaste.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールから表示言語を決める / Pick the UI language from the locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        message: {
            textCreateError: {
                ja: "新規テキスト作成時にエラーが発生しました：\n",
                en: "An error occurred while creating new text:\n"
            },
            replaceError: {
                ja: "テキスト置換中にエラーが発生しました：\n",
                en: "An error occurred while replacing text:\n"
            },
            clipboardError: {
                ja: "クリップボードからの取得に失敗しました：\n",
                en: "Failed to get text from clipboard:\n"
            },
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            }
        }
    };

    /* ラベルをドット区切りのキーで引く / Look a label up by a dot-separated key */
    function L(key) {
        var parts = key.split(".");
        var entry = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!entry) return key;
            entry = entry[parts[i]];
        }
        return (entry && entry[currentLanguage]) ? entry[currentLanguage] : key;
    }

    // =========================================
    // 選択の操作 / Selection helpers
    // =========================================

    /* 選択オブジェクトを配列に写し取る（selection は操作で変化するため）/ Copy the selection into a plain array, since selection changes as we work */
    function captureSelection(doc) {
        var items = [];
        var selection = doc.selection;
        if (!selection || !selection.length) return items;
        for (var i = 0; i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    /* 選択状態を差し替える（失敗しても処理を止めない）/ Replace the selection without aborting the run on failure */
    function setSelection(doc, items) {
        try {
            doc.selection = (items && items.length > 0) ? items : null;
        } catch (e) {
            // Illustrator が選択を拒む場合は現在の選択のままにする / Keep whatever stays selected
        }
    }

    /* 指定のオブジェクトをまとめて削除する / Remove the given objects */
    function removeItems(items) {
        if (!items || !items.length) return;
        for (var i = items.length - 1; i >= 0; i--) {
            try {
                items[i].remove();
            } catch (e) {
                // 既に消えているものは無視する / Ignore items that are already gone
            }
        }
    }

    /* 配列から最初のテキストフレームを探す / Find the first text frame in the list */
    function findTextFrame(items) {
        if (!items) return null;
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "TextFrame") return items[i];
        }
        return null;
    }

    // =========================================
    // クリップボードの取得 / Clipboard access
    // =========================================

    /*
     * 一度ペーストして、貼り付けられたテキストフレームから内容と座標を読み取る。
     * 読み取り後は貼り付けたオブジェクトを削除し、元の選択へ戻す。
     * Paste once, read the contents and bounds from the pasted text frame,
     * then remove the pasted objects and restore the original selection.
     */
    function readClipboardTextFrame(doc, originalSelection) {
        var result = null;
        var pasted = null;

        try {
            app.paste();
            pasted = captureSelection(doc);

            var textFrame = findTextFrame(pasted);
            if (textFrame) {
                result = {
                    bounds: textFrame.geometricBounds,
                    contents: textFrame.contents
                };
            }
        } catch (e) {
            alert(L("message.clipboardError") + e);
        }

        /* 成否にかかわらず、貼り付けた分を消して元の選択へ戻す / Clean up and restore regardless of the outcome */
        removeItems(pasted);
        setSelection(doc, originalSelection);
        return result;
    }

    // =========================================
    // テキストの適用 / Text application
    // =========================================

    /* クリップボードのテキストを元の位置に新規テキストフレームとして作成する / Create a new text frame at the original position */
    function createNewTextFrame(layer, originalBounds, textContent) {
        try {
            var textFrame = layer.textFrames.add();
            textFrame.contents = textContent;

            var bounds = textFrame.geometricBounds;
            textFrame.translate(originalBounds[0] - bounds[0], originalBounds[1] - bounds[1]);
        } catch (e) {
            alert(L("message.textCreateError") + e);
        }
    }

    /*
     * テキストフレームなら内容を置き換え、グループなら中を再帰的にたどる。
     * TextRange が選択されている場合は親のテキストフレームに読み替える。
     * Replace the contents of a text frame, or walk into a group.
     * A selected TextRange is promoted to its parent text frame.
     */
    function applyContentToTarget(target, textContent) {
        try {
            if (target.typename === "TextRange" && target.parent.typename === "TextFrame") {
                target = target.parent;
            }

            if (target.typename === "TextFrame") {
                if (target.kind === TextType.POINTTEXT ||
                    target.kind === TextType.AREATEXT ||
                    target.kind === TextType.PATHTEXT) {
                    target.contents = textContent;
                }
                return;
            }

            if (target.typename === "GroupItem") {
                var items = target.pageItems;
                for (var i = 0; i < items.length; i++) {
                    applyContentToTarget(items[i], textContent);
                }
            }
        } catch (e) {
            alert(L("message.replaceError") + e);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 選択の退避、クリップボードの取得、テキストの適用までを通して行う / Save the selection, read the clipboard, then apply the text */
    function main() {
        if (app.documents.length === 0) {
            alert(L("message.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        var clipboardData = readClipboardTextFrame(doc, originalSelection);
        if (!clipboardData) return;

        if (originalSelection.length === 0) {
            createNewTextFrame(doc.activeLayer, clipboardData.bounds, clipboardData.contents);
        } else {
            for (var i = 0; i < originalSelection.length; i++) {
                applyContentToTarget(originalSelection[i], clipboardData.contents);
            }
        }

        /* 置換後の表示を確実に更新するため、選択を解除してから戻す / Clear and reset the selection so the redraw reflects the change */
        setSelection(doc, null);
        app.redraw();
        setSelection(doc, originalSelection);
    }

    main();

})();
