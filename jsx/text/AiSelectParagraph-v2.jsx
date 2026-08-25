#target illustrator

/*

### 概要

テキストにカーソルを置いた状態で実行すると、その段落全体を選択します。

詳細は README を参照してください。

### Overview

Selects the whole paragraph the text cursor is sitting in.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSelectParagraph-v2";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSelectParagraph-v2.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSelectParagraph-v2.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 段落末尾の改行（\r）を選択範囲に含めるか / Include the trailing return in the selection
 * true: 改行まで選択 / select through the return
 * false: 改行の手前まで選択 / stop before the return
 * 強制改行（）は改行として扱わないため、この設定の影響を受けません
 * A forced line break () is never treated as a return, so it ignores this setting
 * @type {boolean}
 */
var INCLUDE_PARAGRAPH_RETURN = true;

/* 段落区切りの改行コード / The return character that delimits a paragraph */
var PARAGRAPH_RETURN_CHAR = "\r";

/* ===== ローカライズ / Localization ===== */

/* 現在の言語を取得 / Get current language */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義 / Label definitions */
var LABELS = {
    noDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    noTextCursor: {
        ja: "テキストにカーソルを置いてください。",
        en: "Please place the text cursor inside a text object."
    },
    multipleParagraphs: {
        ja: "段落をまたいでいます。1つの段落内にカーソルを置いてください。",
        en: "The selection spans multiple paragraphs. Place the cursor within a single paragraph."
    },
    emptyParagraph: {
        ja: "この段落には選択できる文字がありません。",
        en: "This paragraph has no selectable characters."
    }
};

/**
 * ラベル文字列を取得する / Get a localized label string
 * @param {string} labelKey - LABELS のキー / key of LABELS
 * @returns {string} ローカライズ済み文字列 / localized string
 */
function labelText(labelKey) {
    var labelEntry = LABELS[labelKey];
    return labelEntry ? (labelEntry[currentLanguage] || labelEntry.en) : labelKey;
}

/* ===== 選択の取得 / Selection lookup ===== */

/**
 * テキスト編集中の TextRange を取得する / Get the text range being edited
 * @returns {TextRange|null} テキスト編集中の TextRange、なければ null / the range, or null
 */
function getActiveTextRange() {
    var currentSelection = app.selection;
    /* テキスト編集中のみ app.selection が配列ではなく TextRange になる */
    /* app.selection is a TextRange (not an array) only while editing text */
    return (currentSelection && currentSelection.typename === "TextRange") ? currentSelection : null;
}

/* ===== 段落範囲の算出 / Paragraph range ===== */

/**
 * カーソル位置が属する段落の範囲を求める / Get the range of the paragraph at the cursor
 * @param {string} storyText - ストーリー全文 / full story text
 * @param {TextRange} activeTextRange - 現在のテキスト選択 / current text selection
 * @returns {object} { start: 開始インデックス, end: 排他的な終了インデックス } / { start, exclusive end }
 */
function getParagraphRange(storyText, activeTextRange) {
    /* 直前の改行の次の文字が段落の先頭 / The paragraph starts after the preceding return */
    var paragraphStartIndex = activeTextRange.start <= 0 ? 0 : storyText.lastIndexOf(PARAGRAPH_RETURN_CHAR, activeTextRange.start - 1) + 1;
    var trailingReturnIndex = storyText.indexOf(PARAGRAPH_RETURN_CHAR, activeTextRange.end);

    /* 最終段落には改行が続かない / The last paragraph has no trailing return */
    if (trailingReturnIndex === -1) return { start: paragraphStartIndex, end: storyText.length };

    return {
        start: paragraphStartIndex,
        end: INCLUDE_PARAGRAPH_RETURN ? trailingReturnIndex + 1 : trailingReturnIndex
    };
}

/* ===== メイン処理 / Main ===== */

/**
 * ストーリー内の指定範囲を選択する / Select the given range within a story
 * @param {Story} targetStory - 対象ストーリー / target story
 * @param {number} startIndex - 開始インデックス / start index
 * @param {number} endIndex - 排他的な終了インデックス / exclusive end index
 * @returns {void}
 */
function selectStoryRange(targetStory, startIndex, endIndex) {
    var targetTextRange = targetStory.textRange;
    targetTextRange.start = startIndex;
    targetTextRange.end = endIndex;
    targetTextRange.select();
    app.redraw();
}

/**
 * カーソルのある段落を選択する / Select the paragraph containing the cursor
 * @returns {void}
 */
function selectCurrentParagraph() {
    if (app.documents.length === 0) {
        alert(labelText("noDocument"));
        return;
    }

    var activeTextRange = getActiveTextRange();
    if (!activeTextRange) {
        alert(labelText("noTextCursor"));
        return;
    }

    /* 全文は story.textRange.contents で取得（story.contents は存在しない） */
    /* Read the full text via story.textRange.contents (story.contents does not exist) */
    var targetStory = activeTextRange.story;
    var storyText = targetStory.textRange.contents;

    /* 選択範囲に改行を含む＝段落をまたいでいるため対象を特定できない */
    /* A return inside the selection means it spans paragraphs, so the target is ambiguous */
    var selectedText = storyText.substring(activeTextRange.start, activeTextRange.end);
    if (selectedText.indexOf(PARAGRAPH_RETURN_CHAR) !== -1) {
        alert(labelText("multipleParagraphs"));
        return;
    }

    var paragraphRange = getParagraphRange(storyText, activeTextRange);
    if (paragraphRange.end <= paragraphRange.start) {
        alert(labelText("emptyParagraph"));
        return;
    }

    selectStoryRange(targetStory, paragraphRange.start, paragraphRange.end);
}

selectCurrentParagraph();
