#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストにカーソルを置いた状態で実行すると、その段落全体を選択します。
同じ段落内で文字列を選択している場合も、その段落全体に広げます。

詳細は README を参照してください。

### Overview

Selects the whole paragraph the text cursor is sitting in.
A selection already inside a paragraph is expanded to cover the whole paragraph.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSelectParagraph";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSelectParagraph.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSelectParagraph.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /**
     * 段落末尾の改行（\r）を選択範囲に含めるか / Include the trailing return in the selection
     * true: 改行まで選択。改行だけの空段落も選択できます / select through the return; an empty paragraph can be selected
     * false: 改行の手前まで選択。空段落は選択できません / stop before the return; an empty paragraph cannot be selected
     * 強制改行（\u0003）は改行として扱わないため、この設定の影響を受けません
     * A forced line break (\u0003) is never treated as a return, so it ignores this setting
     * @type {boolean}
     */
    var INCLUDE_PARAGRAPH_RETURN = true;

    /* 段落区切りの改行コード / The return character that delimits a paragraph */
    var PARAGRAPH_RETURN_CHAR = "\r";

    // =========================================
    // ローカライズ / Localization
    // =========================================

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
     * 未定義のキーはキー名をそのまま返す / An unknown key falls back to the key itself
     * @param {string} labelKey - LABELS のキー / key of LABELS
     * @returns {string} ローカライズ済み文字列 / localized string
     */
    function getLocalizedText(labelKey) {
        /* 現在の言語の文言を返し、無ければ英語にフォールバック / Return the string for the current language, falling back to English */
        var labelEntry = LABELS[labelKey];
        return labelEntry ? (labelEntry[currentLanguage] || labelEntry.en) : labelKey;
    }

    // =========================================
    // 選択の取得 / Selection Lookup
    // =========================================

    /**
     * テキスト編集中の TextRange を取得する / Get the text range being edited
     * オブジェクト選択時は配列が返るため、TextRange 以外は対象外とする
     * A selected object comes back as an array, so anything but a TextRange is rejected
     * @returns {TextRange|null} テキスト編集中の TextRange、なければ null / the range, or null
     */
    function getActiveTextRange() {
        var currentSelection = app.selection;
        /* テキスト編集中のみ app.selection が配列ではなく TextRange になる */
        /* app.selection is a TextRange (not an array) only while editing text */
        return (currentSelection && currentSelection.typename === "TextRange") ? currentSelection : null;
    }

    // =========================================
    // 段落範囲の算出 / Paragraph Range
    // =========================================

    /**
     * 指定位置が属する段落の開始インデックスを求める / Get the start index of the paragraph at a position
     * 直前の改行の次の文字が段落の先頭。改行が無ければストーリーの先頭
     * The paragraph starts after the preceding return, or at the story head when there is none
     * @param {string} storyText - ストーリー全文 / full story text
     * @param {number} position - 走査の起点 / position to scan back from
     * @returns {number} 段落の開始インデックス / start index of the paragraph
     */
    function getParagraphStartIndex(storyText, position) {
        /* ストーリー先頭なら探す必要がない / Nothing to scan at the story head */
        if (position <= 0) return 0;

        /* 直前の改行を後方へ探し、その次の文字を段落の先頭とする（見つからなければ -1 + 1 = 0） */
        /* Scan back for the preceding return and take the next character (not found gives -1 + 1 = 0) */
        return storyText.lastIndexOf(PARAGRAPH_RETURN_CHAR, position - 1) + 1;
    }

    /**
     * 選択が段落をまたいでいるか判定する / Check whether the selection spans paragraphs
     * 開始位置と終了位置それぞれの段落先頭を比べる（文字列の切り出しには依存しない）
     * Compares the paragraph head of the start position with that of the end position
     * @param {string} storyText - ストーリー全文 / full story text
     * @param {TextRange} activeTextRange - 現在のテキスト選択 / current text selection
     * @returns {boolean} またいでいれば true / true when it spans paragraphs
     */
    function spansMultipleParagraphs(storyText, activeTextRange) {
        /* 文字列選択中は終了位置の1文字手前で判定する。TextRange.end は排他的なため、 */
        /* 段落末尾の改行まで選んだだけで次段落の先頭とみなされてしまう */
        /* For a selection, judge one character before the end: TextRange.end is exclusive, so */
        /* selecting through the trailing return would otherwise land on the next paragraph's head */
        var endPosition = activeTextRange.end > activeTextRange.start ? activeTextRange.end - 1 : activeTextRange.end;

        /* 両端それぞれの段落先頭を求める / Find the paragraph head for each end of the selection */
        var startParagraphHead = getParagraphStartIndex(storyText, activeTextRange.start);
        var endParagraphHead = getParagraphStartIndex(storyText, endPosition);

        /* 先頭が違えば別の段落にまたがっている / Different heads mean the selection crosses paragraphs */
        return startParagraphHead !== endParagraphHead;
    }

    /**
     * カーソル位置が属する段落の範囲を求める / Get the range of the paragraph at the cursor
     * 段落先頭を求め、そこから最初の改行までを範囲とする。カーソルのみでも文字列選択中でも同じ段落を返す
     * Finds the paragraph head, then runs to the first return after it; a caret and a selection resolve alike
     * 呼び出し前に spansMultipleParagraphs() が同一段落であることを保証している前提
     * Assumes spansMultipleParagraphs() has already confirmed the selection stays in one paragraph
     * @param {string} storyText - ストーリー全文 / full story text
     * @param {TextRange} activeTextRange - 現在のテキスト選択 / current text selection
     * @returns {object} { start: 開始インデックス, end: 排他的な終了インデックス } / { start, exclusive end }
     */
    function getParagraphRange(storyText, activeTextRange) {
        /* 段落先頭から最初の改行を探す。選択の終了位置から探すと、改行まで選択済みのときに */
        /* その改行を飛び越えて次段落の改行を拾ってしまう */
        /* Scan from the paragraph head: scanning from the selection end would skip past a return */
        /* that is already selected and pick up the next paragraph's one */
        var paragraphStartIndex = getParagraphStartIndex(storyText, activeTextRange.start);
        var trailingReturnIndex = storyText.indexOf(PARAGRAPH_RETURN_CHAR, paragraphStartIndex);

        /* 改行が見つからない＝最終段落なのでストーリー末尾までが範囲 */
        /* No return found means the last paragraph, so the range runs to the end of the story */
        if (trailingReturnIndex === -1) return { start: paragraphStartIndex, end: storyText.length };

        /* 設定に応じて末尾の改行を1文字分だけ含める / Extend by one to include the trailing return when set */
        return {
            start: paragraphStartIndex,
            end: INCLUDE_PARAGRAPH_RETURN ? trailingReturnIndex + 1 : trailingReturnIndex
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ストーリー内の指定範囲を選択する / Select the given range within a story
     * story.textRange の start/end を書き換えてから select する
     * Rewrites start/end of story.textRange, then selects it
     * @param {Story} targetStory - 対象ストーリー / target story
     * @param {number} startIndex - 開始インデックス / start index
     * @param {number} endIndex - 排他的な終了インデックス / exclusive end index
     * @returns {void}
     */
    function selectStoryRange(targetStory, startIndex, endIndex) {
        /* ストーリー全体の範囲を取得し、選択したい範囲へ狭める / Take the whole story range and narrow it to the target */
        var targetTextRange = targetStory.textRange;
        targetTextRange.start = startIndex;
        targetTextRange.end = endIndex;

        /* 選択を反映して画面を更新 / Apply the selection and refresh the screen */
        targetTextRange.select();
        app.redraw();
    }

    /**
     * カーソルのある段落を選択する / Select the paragraph containing the cursor
     * ガードを通過したあと段落範囲を求め、空でなければ選択する
     * Derives the paragraph range after the guards, and selects it unless it is empty
     * @returns {void}
     */
    function selectCurrentParagraph() {
        /* ドキュメントが無ければ何もできない / Nothing to do without a document */
        if (app.documents.length === 0) {
            alert(getLocalizedText("noDocument"));
            return;
        }

        /* テキスト編集中でなければ対象の段落が決まらない / Without a text cursor there is no paragraph to act on */
        var activeTextRange = getActiveTextRange();
        if (!activeTextRange) {
            alert(getLocalizedText("noTextCursor"));
            return;
        }

        /* 全文は story.textRange.contents で取得（story.contents は存在しない） */
        /* Read the full text via story.textRange.contents (story.contents does not exist) */
        var targetStory = activeTextRange.story;
        var storyText = targetStory.textRange.contents;

        /* 段落をまたぐ選択は対象を特定できない / A selection spanning paragraphs has no single target */
        if (spansMultipleParagraphs(storyText, activeTextRange)) {
            alert(getLocalizedText("multipleParagraphs"));
            return;
        }

        /* 空段落は INCLUDE_PARAGRAPH_RETURN が false のときだけ選択対象が無くなる */
        /* An empty paragraph has nothing to select only when INCLUDE_PARAGRAPH_RETURN is false */
        var paragraphRange = getParagraphRange(storyText, activeTextRange);
        if (paragraphRange.end <= paragraphRange.start) {
            alert(getLocalizedText("emptyParagraph"));
            return;
        }

        selectStoryRange(targetStory, paragraphRange.start, paragraphRange.end);
    }

    selectCurrentParagraph();

})();
