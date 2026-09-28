#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心に黒く塗った正方形を作成して選択し、「Convert to Shape」「Make Pixel Perfect」コマンドを適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n509eb6aa0a19

### Overview

Creates a black square at the center of the current view, selects it, and applies the
"Convert to Shape" and "Make Pixel Perfect" commands.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewRectangle";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n509eb6aa0a19"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var RECT_SIZE = 100; /* 作成する正方形の一辺 / side length of the square */

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
        alert: {
            noDocument: { ja: "ドキュメントがありません。", en: "No document is open." },
            noEditableLayer: {
                ja: "ロック解除かつ表示されているレイヤーがありません。",
                en: "No unlocked and visible layers found."
            }
        }
    };

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * ドキュメントのカラーモードに応じた黒を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
    function getBlackFillColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var blackCmyk = new CMYKColor();
            blackCmyk.cyan = 0;
            blackCmyk.magenta = 0;
            blackCmyk.yellow = 0;
            blackCmyk.black = 100;
            return blackCmyk;
        }
        var blackRgb = new RGBColor();
        blackRgb.red = 0;
        blackRgb.green = 0;
        blackRgb.blue = 0;
        return blackRgb;
    }

    /**
     * ロックされておらず表示されている最初のレイヤーを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 編集可能なレイヤー。見つからない場合は null
     */
    function getUnlockedVisibleLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (!candidateLayer.locked && candidateLayer.visible) return candidateLayer;
        }
        return null;
    }

    /**
     * 作成先レイヤーを決め、必要ならアクティブレイヤーを切り替える
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 作成先レイヤー。見つからない場合は null
     */
    function resolveTargetLayer(doc) {
        var targetLayer = doc.activeLayer;
        if (!targetLayer.locked && targetLayer.visible) return targetLayer;

        /* ロック／非表示なら、順にロック解除かつ表示のレイヤーを探す / Fall back to the first editable layer */
        var editableLayer = getUnlockedVisibleLayer(doc);
        if (!editableLayer) return null;

        doc.activeLayer = editableLayer;
        return editableLayer;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表示中心に正方形を作成し、選択してコマンドを適用する
     * @returns {void}
     */
    function main() {
        /* ドキュメント確認 / Ensure a document is open */
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        var targetLayer = resolveTargetLayer(doc);
        if (!targetLayer) {
            alert(getLabel(LABELS.alert.noEditableLayer));
            return;
        }

        /* 表示領域の中心座標を取得 / Get view center */
        var viewCenter = doc.activeView.centerPoint;

        /* 正方形を作成 / Create the square */
        var square = targetLayer.pathItems.rectangle(
            viewCenter[1] + RECT_SIZE / 2,
            viewCenter[0] - RECT_SIZE / 2,
            RECT_SIZE,
            RECT_SIZE
        );

        /* カラーモードに応じた黒を設定し、線はなしに / Fill with black, no stroke */
        square.fillColor = getBlackFillColor(doc);
        square.stroked = false;

        /* 作成した正方形だけを選択 / Select the created square only */
        doc.selection = null;
        square.selected = true;

        /* 選択オブジェクトにコマンドを適用 / Apply commands to the selection */
        app.executeMenuCommand("Convert to Shape");
        app.executeMenuCommand("Make Pixel Perfect");
    }

    main();

})();
