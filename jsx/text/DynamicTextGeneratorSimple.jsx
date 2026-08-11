#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した1つのテキストフレームの各行をアウトライン幅で測定し、最長行の幅にそろうよう行ごとの文字サイズを変倍します。あわせて行送りを自動（既定100%）に切り替えます。

詳細はREADMEを参照。

*/

/*

### Overview

Measures each line of a single selected text frame by its outline width and scales every line's font size so all lines match the widest one, then switches leading to auto (100% by default).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DynamicTextGeneratorSimple";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-11";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGeneratorSimple.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGeneratorSimple.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 幅をそろえたあとに行送りを自動へ切り替えるか / Switch leading to auto after fitting the widths */
    var APPLY_AUTO_LEADING = true;

    /* 自動行送りの比率（％）。applyAutoLeading() の既定値でもある
       Auto-leading ratio in percent; also the default used by applyAutoLeading() */
    var AUTO_LEADING_AMOUNT = 100;

    /* 変倍率がこの範囲内なら誤差とみなして変更しない / Treat ratios within this range as no change */
    var RATIO_EPSILON = 0.001;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneTextFrame: { ja: "テキストフレームを1つだけ選択してください。", en: "Select exactly one text frame." },
            needTwoLines: { ja: "2行以上のテキストを選択してください。", en: "Select text with two or more lines." },
            noMeasurableLine: { ja: "有効なテキスト行が見つかりませんでした。", en: "No measurable text line was found." }
        }
    };

    /**
     * LABELS を上から順にたどって現在のUI言語のラベルを返す
     * @param {...string} - LABELS をたどるキー
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 行の測定結果
     * @typedef {object} LineMetric
     * @property {number} start - ストーリー内の開始文字インデックス
     * @property {number} end - ストーリー内の終了文字インデックス（この位置は含まない）
     * @property {number} width - 行の外形幅（pt）。測定できない行は0
     */

    /**
     * 改行や空白しか含まない行かどうかを判定する
     * @param {TextRange} line - 判定する行
     * @returns {boolean} 内容が空とみなせる場合 true
     */
    function isBlankLine(line) {
        return line.contents.replace(/[\r\n\x03\s　]/g, "").length === 0;
    }

    /**
     * 行の内容を一時テキストフレームへ複製し、アウトライン化して外形幅を測る
     * 一時オブジェクトは成否にかかわらず必ず削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextRange} line - 測定する行
     * @returns {number} 行の外形幅（pt）。測定できない場合は0
     */
    function measureLineWidth(doc, line) {
        var tempTextFrame = null;
        var outlineGroup = null;
        var width = 0;

        try {
            tempTextFrame = doc.textFrames.add();
            line.duplicate(tempTextFrame, ElementPlacement.INSIDE);
            /* createOutline() は元のテキストフレームを消費するので参照を手放す
               createOutline() consumes the source frame, so drop the reference */
            outlineGroup = tempTextFrame.createOutline();
            tempTextFrame = null;
            width = outlineGroup.width;
        } catch (e) {
            /* アウトライン化できないときはフレーム幅で代用 / Fall back to the frame width */
            try {
                width = (tempTextFrame !== null) ? tempTextFrame.width : 0;
            } catch (err) {
                width = 0;
            }
        }

        /* 残っている一時オブジェクトを後始末する / Clean up whichever temporary object survived */
        try { if (outlineGroup !== null) outlineGroup.remove(); } catch (e) {}
        try { if (tempTextFrame !== null) tempTextFrame.remove(); } catch (e) {}

        return width;
    }

    /**
     * ストーリー内の指定範囲の文字サイズを変倍する（固定行送りの場合は行送りも追従させる）
     * @param {Story} story - 対象ストーリー
     * @param {number} startIndex - 開始文字インデックス
     * @param {number} endIndex - 終了文字インデックス（この位置は含まない）
     * @param {number} ratio - 変倍率
     * @returns {void}
     */
    function scaleCharacterSizes(story, startIndex, endIndex, ratio) {
        for (var i = startIndex; i < endIndex; i++) {
            try {
                var charAttr = story.characters[i].characterAttributes;
                charAttr.size *= ratio;
                /* 固定行送りのときだけ行が重ならないよう行送りも変倍する
                   Scale leading as well, but only when it is fixed */
                if (!charAttr.autoLeading) {
                    charAttr.leading *= ratio;
                }
            } catch (e) {
                /* 設定できない文字はスキップ / Skip characters that reject the change */
            }
        }
    }

    /**
     * 各行の文字サイズを最長行の幅にそろえる
     * 先に全行を測ってから変倍する（変倍で行が再合成されても対象がずれないようにするため）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function fitLinesToWidestLine(doc, textFrame) {
        var lines = textFrame.lines;
        /** @type {LineMetric[]} */
        var lineMetrics = [];
        var maxWidth = 0;

        /* 1. 全行の外形幅と最大幅を測る（この時点ではテキストを変更しない）
           1. Measure every line and the widest width; the text is not touched yet */
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            var width = isBlankLine(line) ? 0 : measureLineWidth(doc, line);
            lineMetrics.push({ start: line.start, end: line.end, width: width });
            if (width > maxWidth) {
                maxWidth = width;
            }
        }

        if (maxWidth === 0) {
            return false;
        }

        /* 2. 行ごとに変倍する。行オブジェクトではなくストーリー内の文字インデックスで指定して
              変倍による行の再合成の影響を受けないようにする
           2. Scale line by line, addressing characters by story index rather than by line object
              so re-composition during scaling cannot shift the target range */
        var story = textFrame.story;
        for (var i = 0; i < lineMetrics.length; i++) {
            var metric = lineMetrics[i];
            if (metric.width <= 0) continue;

            var ratio = maxWidth / metric.width;
            /* 幅がほぼ同等の行はスキップ / Skip lines that already match the widest one */
            if (Math.abs(ratio - 1) < RATIO_EPSILON) continue;

            scaleCharacterSizes(story, metric.start, metric.end, ratio);
        }

        return true;
    }

    /**
     * テキストフレームの行送りを自動に切り替え、各段落の自動行送り比率を設定する
     * 他のスクリプトからも単体で呼べるよう、対象と比率を引数で受け取る
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} autoLeadingAmount - 自動行送りの比率（％）。省略時は100
     * @returns {void}
     */
    function applyAutoLeading(textFrame, autoLeadingAmount) {
        var amount = (typeof autoLeadingAmount === "number" && !isNaN(autoLeadingAmount)) ? autoLeadingAmount : 100;

        /* フレーム全体を自動行送りにする / Switch the whole frame to auto leading */
        textFrame.textRange.characterAttributes.autoLeading = true;

        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            try {
                paragraphs[i].characterAttributes.autoLeading = true;
                paragraphs[i].paragraphAttributes.autoLeadingAmount = amount;
            } catch (e) {
                /* 空段落など設定できないものはスキップ / Skip paragraphs that reject the setting */
            }
        }
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;

        /* 選択オブジェクトのチェック / Validate the selection */
        if (!selectionItems || selectionItems.length !== 1 || selectionItems[0].typename !== "TextFrame") {
            alert(getLabel("alert", "selectOneTextFrame"));
            return;
        }

        var targetTextFrame = selectionItems[0];
        if (targetTextFrame.lines.length <= 1) {
            alert(getLabel("alert", "needTwoLines"));
            return;
        }

        if (!fitLinesToWidestLine(doc, targetTextFrame)) {
            alert(getLabel("alert", "noMeasurableLine"));
            return;
        }

        if (APPLY_AUTO_LEADING) {
            applyAutoLeading(targetTextFrame, AUTO_LEADING_AMOUNT);
        }

        app.redraw();
    }

    main();

})();
