#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプトの概要：
クリップボードにあるテキストを、選択中のテキストフレームすべてに置換します。
選択がない場合は、新しいテキストフレームを作成し、元の位置に配置されます。
ローカライズ対応済。

処理の流れ：
1. 一度通常のペースト（⌘V / Ctrl+V 相当）を行い、貼り付けられたテキストフレームから内容と位置を取得（貼り付けオブジェクトは削除して元の選択を復元）
2. テキストが取得できなければ終了
3. テキストを選択中オブジェクトに適用（TextFrame または GroupItem 内の TextFrame）
4. 選択がない場合は新規テキストフレームを元の位置に配置

対象：TextFrame、GroupItem 内の TextFrame
非対象：画像、図形、ロックされたオブジェクト

Original Idea by: Gorolib Design

作成日：2024-10-28
更新日：2025-11-24
*/

function getCurrentLang() {
    return ($.locale === "ja" || $.locale === "ja_JP") ? "ja" : "en";
}

var lang = getCurrentLang();

var LABELS = {
    noTextFound: {
        ja: "クリップボードに有効なテキストが見つかりませんでした。",
        en: "No valid text found in clipboard."
    },
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
    }
};

function main() {
    var doc = app.activeDocument;
    doc.undoGroup = "Replace Text With Paste";

    // 選択状態は Illustrator の selection オブジェクトから切り離して退避しておく
    var originalSelection = [];
    var docSel = doc.selection;
    if (docSel && docSel.length && docSel.length > 0) {
        for (var i = 0; i < docSel.length; i++) {
            originalSelection.push(docSel[i]);
        }
    }

    var layer = doc.activeLayer;

    var clipboardData = getClipboardTextFrame(originalSelection);
    if (!clipboardData) {
        // alert(LABELS.noTextFound[lang]);
        return;
    }

    var pastedBounds = clipboardData.bounds;
    var pastedContent = clipboardData.contents;

    if (originalSelection.length === 0) {
        createNewTextFrame(layer, pastedBounds, pastedContent);
    } else {
        for (var i = 0; i < originalSelection.length; i++) {
            applyContentToTarget(originalSelection[i], pastedContent);
        }
    }

    // 差し替え後に一度選択を解除し、再選択してから再描画
    try {
        if (originalSelection && originalSelection.length > 0) {
            // 一度選択解除
            doc.selection = null;
            app.redraw();
            // 元の選択を復元
            doc.selection = originalSelection;
        } else {
            // 選択がなかった場合は単純に再描画のみ
            app.redraw();
        }
    } catch (e) {
        // 選択の復元に失敗しても、スクリプト全体が止まらないようにする
        try {
            app.redraw();
        } catch (e2) {}
    }
}

/*
クリップボードのテキストを元の位置に新規テキストフレームとして作成・配置する
*/
function createNewTextFrame(layer, originalBounds, textContent) {
    try {
        var newTextFrame = layer.textFrames.add();
        newTextFrame.contents = textContent;
        var newBounds = newTextFrame.geometricBounds;

        var deltaX = originalBounds[0] - newBounds[0];
        var deltaY = originalBounds[1] - newBounds[1];
        newTextFrame.translate(deltaX, deltaY);
    } catch (e) {
        alert(LABELS.textCreateError[lang] + e);
    }
}

/*
選択オブジェクトが TextFrame または GroupItem の場合にテキスト内容を適用する
TextRange 選択時は TextFrame に昇格して内容を置換
*/
function applyContentToTarget(target, textContent) {
    try {
        if (target.typename === "TextRange" && target.parent.typename === "TextFrame") {
            target = target.parent;
        }

        if (target.typename === "TextFrame") {
            if (
                target.kind === TextType.POINTTEXT ||
                target.kind === TextType.AREATEXT ||
                target.kind === TextType.PATHTEXT
            ) {
                target.contents = textContent;
            }
        } else if (target.typename === "GroupItem") {
            var items = target.pageItems;
            for (var i = 0; i < items.length; i++) {
                applyContentToTarget(items[i], textContent);
            }
        }
    } catch (e) {
        alert(LABELS.replaceError[lang] + e);
    }
}

/*
クリップボードからテキストフレーム相当の情報を取得する
手順：
3. 一度通常のペースト（app.paste）を実行し、その貼り付けオブジェクトから情報を取得
4. 貼り付けられたオブジェクトは remove() で削除し、Undo は使わずに元の選択状態を復元
戻り値：
  { bounds: [], contents: "" } 形式のオブジェクト
テキストが取得できなければ null
*/
function getClipboardTextFrame(originalSelection) {
    var doc;
    try {
        doc = app.activeDocument;
    } catch (e) {
        return null;
    }

    var result = null;

    try {
        // 3. 一度ペースト（ユーザーの ⌘V / Ctrl+V と同等）
        app.paste();

        // ペースト直後は貼り付けオブジェクトが選択されている前提
        var pasted = doc.selection;

        if (pasted && pasted.length > 0) {
            var textFrame = null;

            // 単一かつ TextFrame の場合
            if (pasted.length === 1 && pasted[0].typename === "TextFrame") {
                textFrame = pasted[0];
            } else {
                // 複数ペーストされた場合は最初に見つかった TextFrame を採用
                for (var i = 0; i < pasted.length; i++) {
                    if (pasted[i].typename === "TextFrame") {
                        textFrame = pasted[i];
                        break;
                    }
                }
            }

            if (textFrame) {
                result = {
                    bounds: textFrame.geometricBounds,
                    contents: textFrame.contents
                };
            }
        }

        // 貼り付けたオブジェクトは削除して元の状態に戻す
        if (pasted && pasted.length > 0) {
            for (var j = pasted.length - 1; j >= 0; j--) {
                try {
                    pasted[j].remove();
                } catch (eRemove) {}
            }
        }

        // 元の選択状態を復元
        try {
            if (originalSelection && originalSelection.length > 0) {
                doc.selection = originalSelection;
            } else {
                doc.selection = null;
            }
        } catch (eRestore) {}

    } catch (e) {
        alert(LABELS.clipboardError[lang] + e);
        return null;
    }

    return result;
}

main();