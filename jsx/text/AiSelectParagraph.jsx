#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

- テキストにカーソルを置いた状態で実行すると、その段落全体を選択します
- 同じ段落内で文字列を選択している場合も、その段落全体に広げます（末尾の改行まで選択済みでも可）
- 段落の区切りは改行（\r）のみです
- 強制改行（Shift+Return, \u0003）は段落内の折り返しとして扱うため、そこで選択は止まりません
- 段落末尾の改行を選択に含めるかは INCLUDE_PARAGRAPH_RETURN で切り換えます

### 処理の流れ

1. テキスト編集中の TextRange を取得
2. ストーリー全文を取得し、選択が段落をまたいでいないか確認
3. カーソル位置から段落先頭を求め、そこから最初の \r までを段落範囲とする
4. story.textRange の start/end を書き換えて選択

### 制限

- テキスト編集中（カーソルがある状態）のみ動作します。オブジェクトを選択しただけでは実行できません
- 段落をまたぐ選択は対象を特定できないため中止します
- 改行だけの空段落は、INCLUDE_PARAGRAPH_RETURN が false のとき選択対象がないため中止します

### 注意

- Illustrator の paragraphs コレクションは強制改行（\u0003）でも区切られるため使用しません
- 全文は story.textRange.contents で取得します（story.contents は存在しません）
- TextRange.end は排他的です。改行まで選択した状態で end をそのまま使うと次段落の先頭を指すため、またぎ判定は end - 1、段落末尾の改行検索は段落先頭を起点にしています

### 更新履歴

- v1.0.1 (2026-07-17): 境界条件を修正。改行まで含めて1段落を選択している場合の誤判定（次段落へはみ出す／複数段落と判定される）を解消
- v1.0.0 (2026-07-16): InDesign 版から移植。段落判定を \r の走査に変更し、強制改行は段落内として扱う

作成日: 2026-07-16
更新日: 2026-07-17
*/

/*
### Overview

- Run with the text cursor placed in text to select the whole paragraph
- A selection within one paragraph is widened to that whole paragraph, even when its trailing return is already selected
- Only a return (\r) delimits a paragraph
- A forced line break (Shift+Return, \u0003) is treated as a wrap inside a paragraph, so the selection does not stop there
- Whether the trailing return is included is controlled by INCLUDE_PARAGRAPH_RETURN

### Flow

1. Get the text range being edited
2. Read the full story text and reject a selection spanning paragraphs
3. Find the paragraph head from the cursor, then run to the first \r after it
4. Select by rewriting start/end of story.textRange

### Limitations

- Works only while editing text; selecting the object itself is not enough
- Aborts when the selection spans paragraphs, as the target is ambiguous
- Aborts on an empty paragraph when INCLUDE_PARAGRAPH_RETURN is false, as nothing is left to select

### Notes

- Illustrator's paragraphs collection also splits on forced line breaks (\u0003), so it is not used here
- The full text comes from story.textRange.contents (story.contents does not exist)
- TextRange.end is exclusive: with the trailing return selected it points at the next paragraph's head, so the span check uses end - 1 and the trailing return is searched from the paragraph head

### Changelog

- v1.0.1 (2026-07-17): Fixed paragraph boundary detection and selection boundary handling; a selection that includes the trailing return no longer spills into the next paragraph or reports as spanning paragraphs
- v1.0.0 (2026-07-16): Ported from the InDesign version; paragraphs are found by scanning for \r, forced line breaks kept inside a paragraph
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.1";

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
