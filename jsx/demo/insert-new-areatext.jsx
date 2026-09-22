#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心に指定サイズの矩形を作り、エリア内文字としてサンプルテキストを流し込み、フォント・サイズ・行送り・行揃えを適用して選択状態にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewAreaText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n509eb6aa0a19

### Overview

Creates a rectangle of a given size at the center of the current view, converts it into an area text
filled with sample text, applies font, size, leading and justification, and leaves it selected.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewAreaText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewAreaText";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewAreaText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewAreaText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n509eb6aa0a19"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var AREA_WIDTH  = 220; /* エリア内文字の幅 / area text width */
    var AREA_HEIGHT = 110; /* エリア内文字の高さ / area text height */

    /* 流し込むサンプルテキスト / Sample text to place */
    var SAMPLE_TEXT = "今日は15:00から予定外のミーティングがありました。疲れを癒すため、夕方に近くのカフェWAVEで、お気に入りの抹茶ラテを楽しみました。\r短い休憩でしたが、心が“ほっと”しました。";

    /* フォント候補（上から順に試す）/ Font candidates, tried in order */
    var FONT_CANDIDATES = [
        "DNPShueiMGoStd-L",
        "HiraginoSans-W3",
        "SourceHanSans-Regular"
    ];

    /* 書式スタイル / Text style */
    var TEXT_STYLE = {
        fontSize: 12,                            /* フォントサイズ（pt）/ font size (pt) */
        leading: 18,                             /* 行送り（pt）/ leading (pt) */
        horizontalScale: 100,                    /* 水平比率（%）/ horizontal scale (%) */
        verticalScale: 100,                      /* 垂直比率（%）/ vertical scale (%) */
        justification: "FULLJUSTIFYLASTLINELEFT" /* 行揃え（Justification のキー名）/ Justification key */
    };

    var APPLY_PROPORTIONAL_METRICS = true; /* プロポーショナルメトリクスを適用するか / apply proportional metrics */

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 表示領域の中心座標を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{x: number, y: number}} 表示中心の座標
     */
    function getViewCenter(doc) {
        var centerPoint = doc.activeView.centerPoint; /* [x, y] */
        return { x: centerPoint[0], y: centerPoint[1] };
    }

    /**
     * ロックされておらず表示されている最初のレイヤーを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 編集可能なレイヤー。見つからない場合は null
     */
    function getFirstEditableLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (!candidateLayer.locked && candidateLayer.visible) return candidateLayer;
        }
        return null;
    }

    /**
     * 指定座標を中心とする矩形を作成する
     * @param {Layer} targetLayer - 作成先レイヤー
     * @param {{x: number, y: number}} centerPoint - 中心座標
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {PathItem} 作成した矩形
     */
    function createCenteredRect(targetLayer, centerPoint, width, height) {
        return targetLayer.pathItems.rectangle(centerPoint.y + height / 2, centerPoint.x - width / 2, width, height);
    }

    /**
     * フォント候補を上から順に試し、最初に見つかったものを適用する
     * @param {CharacterAttributes} characterAttributes - 適用先の文字属性
     * @param {string[]} fontCandidates - フォント名の候補
     * @returns {boolean} 適用できたら true
     */
    function applyFallbackFont(characterAttributes, fontCandidates) {
        for (var i = 0; i < fontCandidates.length; i++) {
            try {
                characterAttributes.textFont = app.textFonts.getByName(fontCandidates[i]);
                return true;
            } catch (e) {
                /* 次の候補を試す / Try the next candidate */
            }
        }
        return false;
    }

    /**
     * テキストフレームに書式スタイルを適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {object} textStyle - TEXT_STYLE と同じ形の書式定義
     * @returns {void}
     */
    function applyTextStyle(textFrame, textStyle) {
        var textRange = textFrame.textRange;
        var characterAttributes = textRange.characterAttributes;

        characterAttributes.size = textStyle.fontSize;
        characterAttributes.leading = textStyle.leading;
        characterAttributes.horizontalScale = textStyle.horizontalScale;
        characterAttributes.verticalScale = textStyle.verticalScale;

        /* 行揃え / Justification */
        textRange.paragraphAttributes.justification = Justification[textStyle.justification];

        /* プロポーショナルメトリクス（未対応バージョンでは無視）/ Proportional metrics (ignored if unsupported) */
        if (APPLY_PROPORTIONAL_METRICS) {
            try {
                characterAttributes.proportionalMetrics = true;
            } catch (e) {
                /* 未対応なら何もしない / Do nothing if unsupported */
            }
        }
    }

    /**
     * 指定オブジェクトだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(doc, pageItem) {
        doc.selection = null;
        doc.selection = [pageItem];
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表示中心にエリア内文字を作成し、書式を適用して選択する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return; /* ドキュメントがなければ終了 / Abort without a document */

        var doc = app.activeDocument;
        var targetLayer = getFirstEditableLayer(doc);
        if (!targetLayer) return; /* すべてロック／非表示なら中断 / Abort if every layer is locked or hidden */

        var viewCenter = getViewCenter(doc);

        /* エリア用の矩形を中心に作成 / Create the area rectangle at the center */
        var areaRect = createCenteredRect(targetLayer, viewCenter, AREA_WIDTH, AREA_HEIGHT);

        /* エリア内文字フレームを作成 / Create the area text frame */
        var textFrame = targetLayer.textFrames.areaText(areaRect);

        /* テキストと体裁を適用 / Apply contents and style */
        textFrame.contents = SAMPLE_TEXT;
        applyTextStyle(textFrame, TEXT_STYLE);
        applyFallbackFont(textFrame.textRange.characterAttributes, FONT_CANDIDATES);

        /* 選択確定 / Finalize selection */
        selectOnly(doc, textFrame);
    }

    main();

})();
