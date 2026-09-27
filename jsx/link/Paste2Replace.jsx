#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した各オブジェクトを、クリップボードの内容で置き換えます。
貼り付けた内容は元のオブジェクトの長辺に合わせて縦横比を保ったまま拡大・縮小し、中心をそろえます。

### 注意

- テキストオブジェクトが選択に含まれる場合は実行しません。
- クリップボードが空の場合は、元のオブジェクトを削除せずに中止します。

### Overview

Replaces each selected object with the contents of the clipboard.
The pasted contents are scaled proportionally to fit the long side of the original object and centered on it.

### Notes

- The script does not run when the selection includes text objects.
- When the clipboard is empty, it stops without deleting the original objects.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "Paste2Replace";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-10-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            alertTitle: { ja: "クリップボードで置き換え", en: "Replace with Clipboard" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "置き換えるオブジェクトを選択してください。", en: "Select the objects to replace." },
            textSelected: {
                ja: "テキストオブジェクトは置き換えられません。\nテキスト以外のオブジェクトだけを選択してください。",
                en: "Text objects cannot be replaced.\nSelect only non-text objects."
            },
            emptyClipboard: {
                ja: "クリップボードが空か、Illustrator に貼り付けられない内容です。",
                en: "The clipboard is empty, or Illustrator cannot paste its contents."
            },
            replaceError: {
                ja: "置き換え中にエラーが発生しました：\n",
                en: "An error occurred while replacing objects:\n"
            }
        }
    };

    /**
     * ラベルをドット区切りのキーで引く
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りキー
     * @returns {string} 現在の表示言語のラベル。見つからない場合はキーをそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        return (labelNode && labelNode[uiLang]) ? labelNode[uiLang] : labelPath;
    }

    /**
     * スクリプト名を見出しにして警告を出す
     * @param {string} message - 表示するメッセージ
     * @returns {void}
     */
    function showAlert(message) {
        alert(message, getLabel("dialog.alertTitle"));
    }

    // =========================================
    // 座標の計算 / Geometry
    // =========================================

    /**
     * visibleBounds から中心と長辺を求める
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {{centerX: number, centerY: number, longSide: number}} 中心座標と長辺の長さ
     */
    function measureBounds(bounds) {
        return {
            centerX: (bounds[0] + bounds[2]) / 2,
            centerY: (bounds[1] + bounds[3]) / 2,
            longSide: Math.max(bounds[2] - bounds[0], bounds[1] - bounds[3])
        };
    }

    /**
     * 複数オブジェクトの visibleBounds を合わせた外接矩形を返す
     * @param {PageItem[]} pageItems - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getUnionBounds(pageItems) {
        var unionBounds = pageItems[0].visibleBounds;
        for (var i = 1; i < pageItems.length; i++) {
            var itemBounds = pageItems[i].visibleBounds;
            unionBounds = [
                Math.min(unionBounds[0], itemBounds[0]),
                Math.max(unionBounds[1], itemBounds[1]),
                Math.max(unionBounds[2], itemBounds[2]),
                Math.min(unionBounds[3], itemBounds[3])
            ];
        }
        return unionBounds;
    }

    // =========================================
    // 置き換え / Replace
    // =========================================

    /**
     * 選択にテキストオブジェクトが含まれるかを調べる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {boolean} 含まれていれば true
     */
    function containsTextObject(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var itemType = selectedItems[i].typename;
            if (itemType === "TextFrame" || itemType === "LegacyTextItem") return true;
        }
        return false;
    }

    /**
     * クリップボードの内容を貼り付け、貼り付いたオブジェクトを返す。
     * 選択を解除してから貼り付けるのは、貼り付けが起きなかったときに
     * 元の選択を「貼り付いたもの」と取り違えないため
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 貼り付いたオブジェクト。何も貼り付かなければ空配列
     */
    function pasteFromClipboard(doc) {
        doc.selection = null;
        app.paste();
        return doc.selection;
    }

    /**
     * 貼り付けたオブジェクト全体を、置き換え先の長辺に合わせて拡大・縮小し、中心をそろえる。
     * 複数ある場合は互いの配置を保ったまま、ひとまとまりとして扱う
     * @param {PageItem[]} pastedItems - 貼り付けたオブジェクト
     * @param {{centerX: number, centerY: number, longSide: number}} targetMetrics - 置き換え先の中心と長辺
     * @returns {void}
     */
    function fitPastedItems(pastedItems, targetMetrics) {
        var pastedMetrics = measureBounds(getUnionBounds(pastedItems));
        var scale = (pastedMetrics.longSide > 0) ? targetMetrics.longSide / pastedMetrics.longSide : 1;

        for (var i = 0; i < pastedItems.length; i++) {
            var pastedItem = pastedItems[i];
            var itemMetrics = measureBounds(pastedItem.visibleBounds);

            /* 各オブジェクトの中心は、全体の中心からの距離を同じ比率で伸縮させた位置へ / Keep each item's offset from the group center, scaled */
            var destX = targetMetrics.centerX + (itemMetrics.centerX - pastedMetrics.centerX) * scale;
            var destY = targetMetrics.centerY + (itemMetrics.centerY - pastedMetrics.centerY) * scale;

            pastedItem.resize(scale * 100, scale * 100);

            var scaledMetrics = measureBounds(pastedItem.visibleBounds);
            pastedItem.translate(destX - scaledMetrics.centerX, destY - scaledMetrics.centerY);
        }
    }

    /**
     * 1つのオブジェクトをクリップボードの内容で置き換える。
     * 貼り付けに成功したときだけ元のオブジェクトを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 置き換えるオブジェクト
     * @returns {boolean} 置き換えたら true、クリップボードが空なら false
     */
    function replaceWithClipboard(doc, targetItem) {
        var targetMetrics = measureBounds(targetItem.visibleBounds);

        var pastedItems = pasteFromClipboard(doc);
        if (pastedItems.length === 0) return false;

        fitPastedItems(pastedItems, targetMetrics);
        targetItem.remove();
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、各オブジェクトをクリップボードの内容で置き換える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            showAlert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        /* 貼り付けで選択が変わるため、先に配列として控える / Keep a copy; pasting changes the selection */
        var targetItems = doc.selection;

        if (!targetItems || targetItems.length === 0) {
            showAlert(getLabel("alert.noSelection"));
            return;
        }
        if (containsTextObject(targetItems)) {
            showAlert(getLabel("alert.textSelected"));
            return;
        }

        try {
            for (var i = targetItems.length - 1; i >= 0; i--) {
                if (!replaceWithClipboard(doc, targetItems[i])) {
                    showAlert(getLabel("alert.emptyClipboard"));
                    return;
                }
            }
        } catch (e) {
            /* ロック中のオブジェクトなど、DOM が操作を拒む場合 / The DOM may refuse, e.g. for locked objects */
            showAlert(getLabel("alert.replaceError") + e.message);
        }
    }

    main();
})();
