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

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            selectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area text." },
            sortError: { ja: "ソート中にエラーが発生しました: ", en: "An error occurred during sorting: " }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.selectAreaText" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en;
    }

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
