#target illustrator

/**
 * AiSelectParagraph.jsx
 *
 * 概要 / Overview:
 * テキストにカーソルを置いた状態で実行すると、その段落全体を選択します。
 * 段落の区切りは改行（\r）のみで、強制改行（Shift+Return, ）は段落内の
 * 折り返しとして扱うため、そこで選択が止まることはありません。
 * 段落末尾の改行を選択に含めるかは INCLUDE_PARAGRAPH_RETURN で切り換えます。
 *
 * Run with the text cursor placed in text to select the whole paragraph.
 * Only a return (\r) delimits a paragraph; a forced line break () is treated
 * as a wrap inside one, so the selection does not stop there. Whether the trailing
 * return is included is controlled by INCLUDE_PARAGRAPH_RETURN.
 *
 * 処理の流れ / Flow:
 * 1. テキスト編集中の TextRange を取得 / Get the text range being edited
 * 2. ストーリー全文を取得し、選択が段落をまたいでいないか確認 / Read the story text, reject a selection spanning paragraphs
 * 3. カーソル位置の前後から \r を探して段落範囲を算出 / Scan for the surrounding returns to derive the range
 * 4. story.textRange の start/end を書き換えて選択 / Select by rewriting start/end of story.textRange
 *
 * 制限 / Limitations:
 * - テキスト編集中（カーソルがある状態）のみ動作します。オブジェクト選択では実行できません
 *   Works only while editing text; selecting the object itself is not enough
 * - 段落をまたぐ選択は対象を特定できないため中止します
 *   Aborts when the selection spans paragraphs, as the target is ambiguous
 *
 * 注意 / Note:
 * Illustrator の paragraphs コレクションは強制改行でも区切られるため使用しません。
 * 全文は story.textRange.contents で取得します（story.contents は存在しません）。
 * Illustrator's paragraphs collection also splits on forced line breaks, so it is not
 * used here; the full text comes from story.textRange.contents (story.contents does not exist).
 *
 * 更新履歴 / Changelog:
 * - v1.0.0 (2026-07-16): InDesign 版から移植。段落判定を \r の走査に変更し、強制改行は段落内として扱う
 *   Ported from the InDesign version; paragraphs are found by scanning for \r, forced line breaks kept inside a paragraph
 *
 * 作成日 / Created: 2026-07-16
 * 更新日 / Updated: 2026-07-16
 */

var SCRIPT_VERSION = "v1.0.0";

/* ===== 設定 / Settings ===== */

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
