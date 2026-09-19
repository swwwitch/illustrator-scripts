#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置に貼り付けます（pasteInPlace）。

詳細は README を参照してください。

### Overview

Pastes the copied objects into every open document at the same position, using Paste in Place.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PasteInAllDocs";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PasteInAllDocs.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PasteInAllDocs.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n04535658c7f6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ローカライズ / Localization
// =========================================
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No documents are open." },
        noCopiedObject: { ja: "コピーされたオブジェクトがありません。", en: "No objects copied." },
        pasteFailed: { ja: "次のドキュメントにはペーストできませんでした：\r{0}", en: "Could not paste into these documents:\r{0}" },
        pasteDone: { ja: "コピーしたオブジェクトをすべてのドキュメントにペーストしました。", en: "Objects pasted into all documents." },
        unexpectedError: { ja: "エラーが発生しました：\r{0}", en: "An error occurred:\r{0}" }
    }
};

var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/**
 * ラベル定義から現在の言語の文言を取得する
 * @param {Object} entry - ja / en を持つラベル定義
 * @returns {string} 現在の言語の文言
 */
function getLabel(entry) {
    return entry[uiLang] || entry.en;
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
// メイン処理 / Main
// =========================================

(function () {

    /**
     * アクティブレイヤーを一時的に編集可能にして処理を実行し、元の状態へ戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Function} action - 編集可能な状態で実行する処理
     * @returns {*} action の戻り値
     */
    function withEditableActiveLayer(doc, action) {
        var layer = doc.activeLayer;
        var wasLocked = layer.locked;
        var wasVisible = layer.visible;
        if (wasLocked) {
            layer.locked = false;
        }
        if (!wasVisible) {
            layer.visible = true;
        }
        try {
            return action();
        } finally {
            if (!wasVisible) {
                layer.visible = false;
            }
            if (wasLocked) {
                layer.locked = true;
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
        // ドキュメントの切り替えが反映される前にメニューコマンドを実行すると、前面のドキュメントにペーストされる
        app.redraw();
        // 「同じ位置にペースト」はアクティブなアートボードが基準になるため、コピー元にそろえる
        if (artboardIndex < doc.artboards.length) {
            doc.artboards.setActiveArtboardIndex(artboardIndex);
        }
    }

    /**
     * 指定ドキュメントへ「同じ位置にペースト」し、追加されたページアイテム数を返す
     * @param {Document} doc - 貼り付け先ドキュメント
     * @param {number} artboardIndex - コピー元のアートボード番号
     * @returns {number} 追加されたページアイテム数
     */
    function pasteInPlaceInto(doc, artboardIndex) {
        activateForPaste(doc, artboardIndex);
        return withEditableActiveLayer(doc, function() {
            // 選択したままペーストすると、直後の操作が選択中のオブジェクトを巻き込む
            doc.selection = null;
            var itemCountBefore = doc.pageItems.length;
            app.executeMenuCommand("pasteInPlace");
            return doc.pageItems.length - itemCountBefore;
        });
    }

    /**
     * コピー元へ試しにペーストして、クリップボードにペーストできるものがあるかを判定する
     * executeMenuCommand() はメニューが無効でも例外を投げないため、ページアイテム数の増減で判定する
     * @param {Document} doc - コピー元ドキュメント
     * @param {number} artboardIndex - コピー元のアートボード番号
     * @returns {boolean} ペーストできるものがあれば true
     */
    function hasPastableContent(doc, artboardIndex) {
        activateForPaste(doc, artboardIndex);
        return withEditableActiveLayer(doc, function() {
            doc.selection = null;
            var itemCountBefore = doc.pageItems.length;
            app.executeMenuCommand("pasteInPlace");
            var pastedItems = doc.selection;
            var addedCount = doc.pageItems.length - itemCountBefore;
            // 試し貼りした複製を削除する（クリップボードはペーストしても保持される）
            for (var i = pastedItems.length - 1; i >= 0; i--) {
                pastedItems[i].remove();
            }
            doc.selection = null;
            return addedCount > 0;
        });
    }

    /**
     * コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置にペーストする
     * @returns {void}
     */
    function main() {
        // ドキュメントが1つも開いていない場合は終了
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        // コピー元のドキュメントとアートボードを保持する
        var sourceDoc = app.activeDocument;
        var sourceArtboardIndex = sourceDoc.artboards.getActiveArtboardIndex();

        try {
            if (!hasPastableContent(sourceDoc, sourceArtboardIndex)) {
                alert(getLabel(LABELS.alert.noCopiedObject));
                return;
            }

            // コピー元以外の各ドキュメントへペースト
            var failedDocNames = [];
            for (var i = 0; i < app.documents.length; i++) {
                var doc = app.documents[i];
                if (doc === sourceDoc) {
                    continue;
                }
                if (pasteInPlaceInto(doc, sourceArtboardIndex) <= 0) {
                    failedDocNames.push(doc.name);
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
            // 元のドキュメントを再度アクティブに
            sourceDoc.activate();
        }
    }

    main();

})();
