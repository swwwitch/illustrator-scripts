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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
