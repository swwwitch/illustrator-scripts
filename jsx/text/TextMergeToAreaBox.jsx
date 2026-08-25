#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

バラバラに分割されたテキストオブジェクトを行単位にまとめ、元のフォント・サイズ・行送り・幅を引き継いだ1つのエリア内文字に再構成します。
PDFをIllustratorで開いたときの分断テキストの復元に使います。

詳細は README を参照してください。

### Overview

Gathers scattered text objects line by line and rebuilds them as a single area text that inherits the original font, size, leading and width.
It is meant for restoring text broken apart when a PDF is opened in Illustrator.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextMergeToAreaBox";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextMergeToAreaBox.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne8d31278c266"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 同じ行と見なすY座標差の閾値（pt） / Y threshold for grouping items into the same line (pt) */
    var LINE_Y_THRESHOLD = 5;

    /* 行送りの最小倍率（フォントサイズ基準） / Minimum leading ratio against the font size */
    var MIN_LEADING_RATIO = 1.2;

    /* エリア内文字に適用する禁則 / Kinsoku applied to the area text */
    var DEFAULT_KINSOKU  = "Soft_v2"; /* 弱い禁則 v2 / Loose v2 */
    var FALLBACK_KINSOKU = "Soft";    /* v2がないバージョン向け / For versions without v2 */

    /* エリア内文字に適用する行揃え / Justification applied to the area text */
    var AREA_TEXT_JUSTIFICATION = Justification.FULLJUSTIFYLASTLINELEFT; /* 両端揃え（最終行左揃え） / Justify, last line left */

    /* あふれ解消で枠を下に伸ばす最大回数 / Maximum number of times the frame is extended downward to clear overflow */
    var MAX_HEIGHT_GROW_STEPS = 20;

    // =========================================
    // ローカライズ / Localize
    // =========================================
    var currentLanguage = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* ラベル定義 / Label definitions */
    var LABELS = {
        alert: {
            noText: { ja: "変換できるテキストはありません。", en: "No convertible text found." }
        }
    };

    /**
     * "category.key" 形式のラベルを現在の言語で取得する
     * @param {string} key - ラベルキー（例: "alert.noText"）
     * @returns {string} 現在の言語のラベル文字列（未定義のときはキーをそのまま返す）
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode) {
            if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
            if (typeof labelNode.en === "string") return labelNode.en;
        }
        return key;
    }

    /**
     * テキストフレームをまとめて削除する
     * @param {Array<TextFrame>} framesToRemove - 削除するテキストフレーム
     * @returns {void}
     */
    function removeFrames(framesToRemove) {
        for (var i = 0; i < framesToRemove.length; i++) {
            framesToRemove[i].remove();
        }
    }

    /**
     * テキストフレームを左から右へ並び替える
     * @param {Array<TextFrame>} framesToSort - 対象テキストフレーム
     * @returns {Array<TextFrame>} 並び替えた新しい配列
     */
    function sortFramesLeftToRight(framesToSort) {
        return framesToSort.slice(0).sort(function(frameA, frameB) {
            return frameA.left - frameB.left;
        });
    }

    /**
     * テキストフレームを上から下へ並び替える
     * @param {Array<TextFrame>} framesToSort - 対象テキストフレーム
     * @returns {Array<TextFrame>} 並び替えた新しい配列
     */
    function sortFramesTopToBottom(framesToSort) {
        return framesToSort.slice(0).sort(function(frameA, frameB) {
            return frameB.position[1] - frameA.position[1];
        });
    }

    /**
     * 上から下に並べたフレームを、Y位置が近いものごとに同じ行としてまとめる
     * @param {Array<TextFrame>} framesTopToBottom - 上から下へ並び替え済みのテキストフレーム
     * @param {number} yThreshold - 同じ行と見なすY座標差の閾値（pt）
     * @returns {Array<Array<TextFrame>>} 行ごとのテキストフレーム配列
     */
    function groupFramesIntoLines(framesTopToBottom, yThreshold) {
        var lineGroups = [];
        for (var i = 0; i < framesTopToBottom.length; i++) {
            var frame = framesTopToBottom[i];
            /* 上から順に並んでいるので、直前の行とだけ比較すればよい
               The frames are sorted top-down, so comparing with the last line is enough */
            var lastLine = lineGroups[lineGroups.length - 1];
            if (lastLine && Math.abs(lastLine[0].position[1] - frame.position[1]) <= yThreshold) {
                lastLine.push(frame);
            } else {
                lineGroups.push([frame]);
            }
        }
        return lineGroups;
    }

    /**
     * 選択からテキストフレームだけを集め、Y位置で行単位に分類する
     * @param {Array<PageItem>} selectedItems - 選択中のオブジェクト
     * @returns {Array<Array<TextFrame>>} 上から順に並べた行ごとのテキストフレーム配列
     */
    function groupTextFramesByLine(selectedItems) {
        var textFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") {
                textFrames.push(selectedItems[i]);
            }
        }
        return groupFramesIntoLines(sortFramesTopToBottom(textFrames), LINE_Y_THRESHOLD);
    }

    /**
     * 1行分のテキストフレームをX順に連結し、1つのテキストフレームにまとめる
     * @param {Array<TextFrame>} lineFrames - 同じ行に属するテキストフレーム
     * @returns {TextFrame} 連結後のテキストフレーム（元のフレームは削除される）
     */
    function mergeLineFrames(lineFrames) {
        var framesInLine = sortFramesLeftToRight(lineFrames);

        var mergedText = "";
        for (var i = 0; i < framesInLine.length; i++) {
            mergedText += framesInLine[i].contents;
        }

        var leftmostFrame = framesInLine[0];
        var originalPosition = leftmostFrame.position;

        var mergedFrame = leftmostFrame.duplicate();
        mergedFrame.orientation = TextOrientation.HORIZONTAL;
        mergedFrame.move(leftmostFrame, ElementPlacement.PLACEBEFORE);
        mergedFrame.contents = mergedText;
        mergedFrame.position = originalPosition;

        removeFrames(framesInLine);
        return mergedFrame;
    }

    /**
     * 各行のテキストを1つの文字列に連結する
     * 文末が「。」「！」「？」または半角の「.」「!」「?」のときだけ改行し、
     * 英単語どうしが隣り合う場合はスペースを挿入、ハイフンで分断されている場合はハイフンを除去して結合する
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム
     * @returns {string} エリア内文字に流し込む文字列
     */
    function joinLineContents(mergedLineFrames) {
        var joinedText = "";
        for (var i = 0; i < mergedLineFrames.length; i++) {
            var lineText = mergedLineFrames[i].contents;
            joinedText += lineText;

            /* 最終行は連結処理なし / No joining after the last line */
            if (i >= mergedLineFrames.length - 1) {
                continue;
            }

            var nextLineText = mergedLineFrames[i + 1].contents;
            var endsWithENWord = /[A-Za-z0-9)]$/.test(lineText);
            var startsWithENWord = /^[A-Za-z0-9(]/.test(nextLineText);

            if (startsWithENWord) {
                if (endsWithENWord) {
                    /* 英単語どうしの間にスペース / Insert a space between English words */
                    joinedText += " ";
                } else if (/[A-Za-z0-9)]-$/.test(lineText)) {
                    /* ハイフン分断を結合 / Join a hyphen-broken word */
                    joinedText = joinedText.replace(/-$/, "");
                }
            }

            /* 文末でのみ改行を追加（和文・欧文とも文末の約物で判定）
               Break only at sentence ends (both Japanese and Western punctuation) */
            if (/[。！？.!?]$/.test(lineText)) {
                joinedText += "\r";
            }
        }
        return joinedText;
    }

    /**
     * 隣り合う2行のY差から行送りを求める
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム
     * @param {number} fontSize - 基準になるフォントサイズ（pt）
     * @returns {number|null} 行送り（pt）。2行未満のときは null
     */
    function computeLeading(mergedLineFrames, fontSize) {
        if (mergedLineFrames.length < 2) {
            return null;
        }
        var firstLineTop = mergedLineFrames[0].position[1];
        var secondLineTop = mergedLineFrames[1].position[1];
        /* 行送りとしてY差を使用 / Use the Y gap as the leading */
        var leading = Math.abs(firstLineTop - secondLineTop);
        if (leading < fontSize) {
            /* 最小でも MIN_LEADING_RATIO 倍にする / Keep at least MIN_LEADING_RATIO times the font size */
            leading = fontSize * MIN_LEADING_RATIO;
        }
        return leading;
    }

    /**
     * 複数フレーム全体のバウンディングボックスを取得する
     * @param {Array<TextFrame>} sourceFrames - 対象テキストフレーム（1つ以上）
     * @returns {Array<number>} [左, 上, 右, 下] の座標（pt）
     */
    function getCombinedVisibleBounds(sourceFrames) {
        var minX = sourceFrames[0].visibleBounds[0];
        var maxY = sourceFrames[0].visibleBounds[1];
        var maxX = sourceFrames[0].visibleBounds[2];
        var minY = sourceFrames[0].visibleBounds[3];
        for (var i = 1; i < sourceFrames.length; i++) {
            var bounds = sourceFrames[i].visibleBounds;
            if (bounds[0] < minX) minX = bounds[0];
            if (bounds[1] > maxY) maxY = bounds[1];
            if (bounds[2] > maxX) maxX = bounds[2];
            if (bounds[3] < minY) minY = bounds[3];
        }
        return [minX, maxY, maxX, minY];
    }

    /**
     * 禁則を適用する（指定の禁則名が使えないバージョンではフォールバックに切り替える）
     * @param {TextRange} targetRange - 適用先のテキスト範囲
     * @param {string} kinsokuName - 適用したい禁則名
     * @param {string} fallbackKinsokuName - 使えなかったときに適用する禁則名
     * @returns {void}
     */
    function applyKinsoku(targetRange, kinsokuName, fallbackKinsokuName) {
        try {
            targetRange.paragraphAttributes.kinsoku = kinsokuName;
            /* 例外にならず無視される場合もあるため読み戻して確認
               Read the value back, since an unknown name can be ignored instead of throwing */
            if (targetRange.paragraphAttributes.kinsoku === kinsokuName) {
                return;
            }
        } catch (e) {
            /* 未対応の禁則名 / Unsupported kinsoku name */
        }
        targetRange.paragraphAttributes.kinsoku = fallbackKinsokuName;
    }

    /**
     * エリア内文字からテキストがあふれているかを判定する
     * @param {TextFrame} areaTextFrame - 判定するエリア内文字
     * @returns {boolean} あふれていれば true
     */
    function isTextOverflowing(areaTextFrame) {
        var composedCharacters = 0;
        for (var i = 0; i < areaTextFrame.lines.length; i++) {
            composedCharacters += areaTextFrame.lines[i].characters.length;
        }
        /* 改行コードは各行の文字数に含まれないため、比較対象から除く
           Line breaks are not counted in each line, so exclude them from the total */
        return composedCharacters < areaTextFrame.contents.replace(/[\r\n]/g, "").length;
    }

    /**
     * テキストがあふれなくなるまで枠を下方向に伸ばす（上端は動かさない）
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @param {number} heightStep - 1回あたりに伸ばす高さ（pt）
     * @returns {void}
     */
    function growFrameUntilTextFits(areaTextFrame, heightStep) {
        var frameTop = areaTextFrame.top;
        for (var i = 0; i < MAX_HEIGHT_GROW_STEPS && isTextOverflowing(areaTextFrame); i++) {
            areaTextFrame.height = areaTextFrame.height + heightStep;
            /* 上端は元の位置に固定 / Keep the top edge where it was */
            areaTextFrame.top = frameTop;
        }
    }

    /**
     * エリア内文字にフォント・サイズ・行揃え・禁則・行送りを適用する
     * @param {TextFrame} areaTextFrame - 適用先のエリア内文字
     * @param {TextFrame} sourceFrame - 書式の引き継ぎ元（一番上の行）
     * @param {number} fontSize - 適用するフォントサイズ（pt）
     * @param {number|null} leading - 適用する行送り（pt）。null なら自動行送りのまま
     * @returns {void}
     */
    function applyAreaTextFormatting(areaTextFrame, sourceFrame, fontSize, leading) {
        var targetRange = areaTextFrame.textRange;
        targetRange.characterAttributes.textFont = sourceFrame.textRange.characterAttributes.textFont;
        targetRange.characterAttributes.size = fontSize;

        /* 禁則・行揃えはテキスト全体に適用する（paragraphs[0] だけでは［段落］パネルに反映されない）
           Apply kinsoku and justification to the whole text (paragraphs[0] alone is not reflected in the Paragraph panel) */
        applyKinsoku(targetRange, DEFAULT_KINSOKU, FALLBACK_KINSOKU);
        targetRange.paragraphAttributes.justification = AREA_TEXT_JUSTIFICATION;

        if (leading !== null) {
            targetRange.characterAttributes.autoLeading = false;
            targetRange.characterAttributes.leading = leading;
        }
    }

    /**
     * 行ごとのフレーム全体の外接矩形からエリア内文字を作成し、体裁を引き継ぐ
     * @param {Array<TextFrame>} mergedLineFrames - 行ごとに連結済みのテキストフレーム（2行以上）
     * @param {string} joinedText - 流し込む文字列
     * @returns {TextFrame} 作成したエリア内文字
     */
    function createAreaTextFrame(mergedLineFrames, joinedText) {
        var fontSize = mergedLineFrames[0].textRange.characterAttributes.size;

        /* 選択状態に依存しないよう mergedLineFrames から外接矩形を取得
           Read the bounding box from mergedLineFrames so it does not depend on the current selection */
        var lineBounds = getCombinedVisibleBounds(mergedLineFrames);
        var boundsLeft = lineBounds[0];
        var boundsTop = lineBounds[1];
        var boundsWidth = lineBounds[2] - boundsLeft;
        var areaHeight = boundsTop - lineBounds[3];

        /* 作成幅は1文字分縮める。縮めると狭くなりすぎる選択では最低1文字分を確保
           Shrink the created width by one character, but keep at least one character for narrow selections */
        var areaWidth = boundsWidth - fontSize;
        if (areaWidth < fontSize) {
            areaWidth = Math.max(boundsWidth, fontSize);
        }

        /* 長方形を作成してエリア内文字に変換 / Create a rectangle and convert it to area text */
        var areaRect = activeDocument.pathItems.rectangle(boundsTop, boundsLeft, areaWidth, areaHeight);
        areaRect.stroked = false;
        areaRect.filled = false;

        var leading = computeLeading(mergedLineFrames, fontSize);
        var areaTextFrame = activeDocument.textFrames.areaText(areaRect);
        areaTextFrame.contents = joinedText;
        applyAreaTextFormatting(areaTextFrame, mergedLineFrames[0], fontSize, leading);

        /* 枠の高さは元の外接矩形どおりなので、最終行があふれる分だけ下に伸ばす
           The frame keeps the original bounding box height, so extend it downward until the last line fits */
        growFrameUntilTextFits(areaTextFrame, (leading !== null) ? leading : fontSize * MIN_LEADING_RATIO);
        return areaTextFrame;
    }

    /**
     * メイン処理：選択テキストを行ごとに連結し、エリア内文字を生成する
     * @returns {void}
     */
    function main() {
        /* ドキュメント未オープン時は終了 / Exit when no document is open */
        if (app.documents.length === 0) {
            return;
        }

        var lineGroups = groupTextFramesByLine(activeDocument.selection);
        if (lineGroups.length === 0) {
            /* エラーメッセージの表示 / Show error message */
            alert(getLabel("alert.noText"));
            return;
        }

        /* 1行だけの場合は別処理：エリア内文字にせず左揃えで出力
           Single line: merge and left-align only, without converting to area text */
        if (lineGroups.length === 1) {
            var singleLineFrame = mergeLineFrames(lineGroups[0]);
            singleLineFrame.textRange.paragraphAttributes.justification = Justification.LEFT;
            app.selection = [singleLineFrame];
            return;
        }

        /* 各行を1つのテキストフレームに連結 / Merge each line into a single text frame */
        var mergedLineFrames = [];
        for (var i = 0; i < lineGroups.length; i++) {
            mergedLineFrames.push(mergeLineFrames(lineGroups[i]));
        }

        var areaTextFrame = createAreaTextFrame(mergedLineFrames, joinLineContents(mergedLineFrames));

        /* 連結に使った元のフレームを削除 / Delete the merged source frames */
        removeFrames(mergedLineFrames);

        /* 生成されたエリア内文字を選択状態にする / Select the generated area text */
        app.selection = null;
        app.selection = [areaTextFrame];
        app.redraw();
    }

    main();
})();
