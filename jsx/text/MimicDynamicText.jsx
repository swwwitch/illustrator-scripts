#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイナミックテキストのような見た目を、通常のテキストで再現します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MimicDynamicText.md

### Overview

Reproduces the look of dynamic text using ordinary text objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MimicDynamicText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MimicDynamicText";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MimicDynamicText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MimicDynamicText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            selectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area text." },
            sortError: { ja: "ソート中にエラーが発生しました: ", en: "An error occurred during sorting: " }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エリア内文字を行ごとのポイント文字に分け、行の幅を元の幅にそろえてから1つのエリア内文字にまとめ直す
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {void}
     */
    function splitTextFrameIntoLines(areaTextFrame) {
        var doc = app.activeDocument;
        var lineTexts = areaTextFrame.contents.split('\r');
        var originalPosition = areaTextFrame.position;
        var currentY = originalPosition[1];
        var sourceAttributes = areaTextFrame.textRange.characterAttributes;
        var textSize = sourceAttributes.size;
        var textFont = sourceAttributes.textFont;
        var textColor = sourceAttributes.fillColor;

        var lineFrames = [];

        for (var i = 0; i < lineTexts.length; i++) {
            var lineFrame = doc.textFrames.pointText([originalPosition[0], currentY]);
            lineFrame.contents = lineTexts[i];
            /* 内容を入れてから範囲を取る / Take the range after setting the contents */
            var lineAttributes = lineFrame.textRange.characterAttributes;
            lineAttributes.size = textSize;
            lineAttributes.textFont = textFont;
            lineAttributes.fillColor = textColor;

            /* 行の幅を元のフレームの幅に合わせる / Scale the line to the original frame width */
            var widthRatio = areaTextFrame.width / lineFrame.width;
            if (widthRatio > 0) {
                lineAttributes.size = textSize * widthRatio;
                lineAttributes.horizontalScale = 100;
                lineAttributes.verticalScale = 100;
            }

            currentY -= lineFrame.height;
            lineFrames.push(lineFrame);
        }
        var mergedFrame = mergeTextFramesVertically(lineFrames);
        if (mergedFrame) {
            /* 行送りを自動にし、自動行送りの値を 110% にする / Set auto leading and autoLeadingAmount to 110 */
            mergedFrame.textRange.characterAttributes.autoLeading = true;

            var mergedParagraphs = mergedFrame.paragraphs;
            for (var j = 0; j < mergedParagraphs.length; j++) {
                mergedParagraphs[j].paragraphAttributes.autoLeadingAmount = 110;
            }

            redraw();
            areaTextFrame.remove();
            mergedFrame.convertPointObjectToAreaObject();
        }
    }

    /**
     * テキストフレームのときだけリストに追加する
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @param {TextFrame[]} textFrameList - 追加先のリスト
     * @returns {void}
     */
    function collectTextFrame(pageItem, textFrameList) {
        if (pageItem.typename === 'TextFrame') {
            textFrameList.push(pageItem);
        }
    }

    /**
     * 複数のテキストフレームを上から順に1つのテキストフレームへ連結する
     * @param {TextFrame[]} textFrames - 連結するテキストフレーム
     * @returns {TextFrame|undefined} 連結後のテキストフレーム（2つ未満のときは undefined）
     */
    function mergeTextFramesVertically(textFrames) {
        if (textFrames.length < 2) {
            return;
        }

        /* 複数行のフレームは1行ずつの複製に分ける / Split multi-line frames into one duplicate per line */
        var sortedFrames = sortTextFramesByPosition(textFrames);
        var singleLineFrames = [];
        for (var i = 0; i < sortedFrames.length; i++) {
            var lineTexts = sortedFrames[i].contents.split('\r');
            for (var j = 0; j < lineTexts.length; j++) {
                if (lineTexts[j] !== "") {
                    var lineFrame = sortedFrames[i].duplicate();
                    lineFrame.contents = lineTexts[j];
                    lineFrame.top -= j * 2000; /* 行順を維持するために位置調整 / Adjust position to maintain line order */
                    singleLineFrames.push(lineFrame);
                }
            }
            sortedFrames[i].remove();
        }
        sortedFrames = sortTextFramesByPosition(singleLineFrames);

        var baseFrame = sortedFrames[0];
        for (var k = 1; k < sortedFrames.length; k++) {
            baseFrame.paragraphs.add('\n');
            var sourceParagraphs = sortedFrames[k].paragraphs;
            for (j = 0; j < sourceParagraphs.length; j++) {
                sourceParagraphs[j].duplicate(baseFrame);
            }
            sortedFrames[k].remove();
        }
        return baseFrame;
    }

    /**
     * テキストフレームを位置（上→下、同じ高さなら左→右）で並べ替えた配列を返す
     * @param {TextFrame[]} frameList - 並べ替えるテキストフレーム
     * @returns {TextFrame[]} 並べ替えた新しい配列（失敗したときは元の配列）
     */
    function sortTextFramesByPosition(frameList) {
        /* 位置の読み取りは DOM アクセス / Reading positions touches the DOM */
        try {
            var sortedList = [];
            for (var i = 0; i < frameList.length; i++) {
                sortedList.push(frameList[i]);
            }

            sortedList.sort(function (firstFrame, secondFrame) {
                var firstPosition = firstFrame.position;
                var secondPosition = secondFrame.position;
                if (firstPosition[1] > secondPosition[1]) {
                    return -1;
                }
                if (firstPosition[1] < secondPosition[1]) {
                    return 1;
                }
                if (firstPosition[1] === secondPosition[1]) {
                    if (firstPosition[0] < secondPosition[0]) {
                        return -1;
                    }
                    if (firstPosition[0] > secondPosition[0]) {
                        return 1;
                    }
                    return 0;
                }
            });
            return sortedList;
        } catch (e) {
            alert(getLabel("alert.sortError") + e.message);
            return frameList;
        }
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        /* 選択確認 / Check selection */
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        var doc = app.activeDocument;
        var areaTextFrame = doc.selection[0];

        if (areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        /* 分割 / Split */
        splitTextFrameIntoLines(areaTextFrame);

        /* 連結 / Merge */
        var selectedItems = doc.selection;
        if (selectedItems.length >= 2) {
            var textFrames = [];
            for (var i = 0; i < selectedItems.length; i++) {
                collectTextFrame(selectedItems[i], textFrames);
            }
            if (textFrames.length >= 2) {
                mergeTextFramesVertically(textFrames);
            }
        }
    }

    main();

})();
