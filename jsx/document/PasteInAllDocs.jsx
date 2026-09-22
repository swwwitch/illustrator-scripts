#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置に貼り付けます（pasteInPlace）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PasteInAllDocs.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n04535658c7f6

### Overview

Pastes the copied objects into every open document at the same position, using Paste in Place.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PasteInAllDocs.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PasteInAllDocs";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PasteInAllDocs.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PasteInAllDocs.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n04535658c7f6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No documents are open." },
            noCopiedObject: { ja: "コピーされたオブジェクトがありません。", en: "No objects copied." },
            pasteFailed: {
                ja: "次のドキュメントにはペーストできませんでした：\r{0}",
                en: "Could not paste into these documents:\r{0}"
            },
            pasteDone: { ja: "コピーしたオブジェクトをすべてのドキュメントにペーストしました。", en: "Objects pasted into all documents." },
            unexpectedError: {
                ja: "エラーが発生しました：\r{0}",
                en: "An error occurred:\r{0}"
            }
        }
    };

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    /**
     * ラベル定義から現在の言語の文言を取得する
     * @param {Object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 現在の言語の文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * 文言内の {0} を指定した文字列に差し替える
     * @param {string} template - {0} を含む文言
     * @param {string} value - 差し込む文字列
     * @returns {string} 差し込み後の文言
     */
    function formatLabel(template, value) {
        return template.split("{0}").join(String(value));
    }

    // =========================================
    // ペースト / Paste
    // =========================================

    /**
     * アクティブレイヤーを一時的に編集可能にして処理を実行し、元の状態へ戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Function} action - 編集可能な状態で実行する処理
     * @returns {*} action の戻り値
     */
    function withEditableActiveLayer(doc, action) {
        var activeLayer = doc.activeLayer;
        var wasLocked = activeLayer.locked;
        var wasVisible = activeLayer.visible;
        if (wasLocked) {
            activeLayer.locked = false;
        }
        if (!wasVisible) {
            activeLayer.visible = true;
        }
        /* 処理が失敗してもレイヤーの状態は戻す / Restore the layer state even if the action fails */
        try {
            return action();
        } finally {
            if (!wasVisible) {
                activeLayer.visible = false;
            }
            if (wasLocked) {
                activeLayer.locked = true;
            }
        }
    }

    /**
     * ドキュメントをアクティブにして、コピー元と同じアートボードを選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {number} artboardIndex - コピー元のアートボード番号
     * @returns {void}
     */
    function activateForPaste(doc, artboardIndex) {
        doc.activate();
        /* 切り替えが反映される前にコマンドを実行すると前面のドキュメントにペーストされる / Redraw so the paste goes to this document */
        app.redraw();
        /* 「同じ位置にペースト」はアクティブなアートボードが基準なのでコピー元にそろえる / Paste in Place follows the active artboard */
        if (artboardIndex < doc.artboards.length) {
            doc.artboards.setActiveArtboardIndex(artboardIndex);
        }
    }

    /**
     * 選択を解除してから「同じ位置にペースト」を実行し、追加されたページアイテム数を返す
     * executeMenuCommand() はメニューが無効でも例外を投げないため、ページアイテム数の増減で判定する
     * @param {Document} doc - 貼り付け先ドキュメント（アクティブであること）
     * @returns {number} 追加されたページアイテム数
     */
    function pasteInPlaceAndCount(doc) {
        /* 選択したままペーストすると直後の操作が選択中のオブジェクトを巻き込む / Deselect so later steps touch only the pasted items */
        doc.selection = null;
        var itemCountBefore = doc.pageItems.length;
        app.executeMenuCommand("pasteInPlace");
        return doc.pageItems.length - itemCountBefore;
    }

    /**
     * 指定ドキュメントへ「同じ位置にペースト」し、追加されたページアイテム数を返す
     * @param {Document} doc - 貼り付け先ドキュメント
     * @param {number} artboardIndex - コピー元のアートボード番号
     * @returns {number} 追加されたページアイテム数
     */
    function pasteInPlaceInto(doc, artboardIndex) {
        activateForPaste(doc, artboardIndex);
        return withEditableActiveLayer(doc, function () {
            return pasteInPlaceAndCount(doc);
        });
    }

    /**
     * コピー元へ試しにペーストして、クリップボードにペーストできるものがあるかを判定する
     * @param {Document} doc - コピー元ドキュメント
     * @param {number} artboardIndex - コピー元のアートボード番号
     * @returns {boolean} ペーストできるものがあれば true
     */
    function hasPastableContent(doc, artboardIndex) {
        activateForPaste(doc, artboardIndex);
        return withEditableActiveLayer(doc, function () {
            var addedCount = pasteInPlaceAndCount(doc);
            var pastedItems = doc.selection;
            /* 試し貼りした複製を削除する（クリップボードはペーストしても保持される）/ Remove the trial paste; the clipboard is kept */
            for (var i = pastedItems.length - 1; i >= 0; i--) {
                pastedItems[i].remove();
            }
            doc.selection = null;
            return addedCount > 0;
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置にペーストする
     * @returns {void}
     */
    function main() {
        /* ドキュメントが1つも開いていない場合は終了 / Abort when no document is open */
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        /* コピー元のドキュメントとアートボードを保持する / Remember the source document and artboard */
        var sourceDoc = app.activeDocument;
        var sourceArtboardIndex = sourceDoc.artboards.getActiveArtboardIndex();

        /* 途中の例外は内容を表示し、最後にコピー元へ戻す / Report errors and always return to the source document */
        try {
            if (!hasPastableContent(sourceDoc, sourceArtboardIndex)) {
                alert(getLabel(LABELS.alert.noCopiedObject));
                return;
            }

            /* コピー元以外の各ドキュメントへペースト / Paste into every other document */
            var failedDocNames = [];
            for (var i = 0; i < app.documents.length; i++) {
                var targetDoc = app.documents[i];
                if (targetDoc === sourceDoc) {
                    continue;
                }
                if (pasteInPlaceInto(targetDoc, sourceArtboardIndex) <= 0) {
                    failedDocNames.push(targetDoc.name);
                }
            }

            if (failedDocNames.length > 0) {
                alert(formatLabel(getLabel(LABELS.alert.pasteFailed), failedDocNames.join("\r")));
            } else {
                alert(getLabel(LABELS.alert.pasteDone));
            }

        } catch (err) {
            alert(formatLabel(getLabel(LABELS.alert.unexpectedError), err));
        } finally {
            /* 元のドキュメントを再度アクティブに / Reactivate the source document */
            sourceDoc.activate();
        }
    }

    main();

})();
