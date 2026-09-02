#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カーソルのある段落を、ひとつ下の段落と入れ替えます。
sky-chaser-high 氏の moveLineDown.jsx（Visual Studio Code の「行を下へ移動」相当）を、
表示行ではなく段落単位で動かすように改変したものです。

詳細は README を参照してください。

### Overview

Swaps the paragraph containing the cursor with the paragraph below it.
A paragraph-based variant of moveLineDown.jsx by sky-chaser-high,
which reproduces Visual Studio Code's "Move Line Down".

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "moveParagraphDown";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "sky-chaser-high";              /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-08-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/moveParagraphDown.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/moveParagraphDown.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 原作 / Original work
 * @author sky-chaser-high
 * @discussion https://github.com/sky-chaser-high/adobe-illustrator-scripts/blob/main/README_ja.md#%E8%A1%8C%E3%82%92%E4%B8%8A--%E4%B8%8B%E3%81%B8%E7%A7%BB%E5%8B%95
 */

(function () {
    /* ドキュメントが開いていない / No document is open */
    if (!app.documents.length) return;

    /* 文字ツールでテキストを選択、またはテキスト内にカーソルがあるときだけ実行する */
    /* Run only when text is selected or the caret is inside text */
    var selectedTextRange = app.activeDocument.selection;
    if (!selectedTextRange || selectedTextRange.typename != 'TextRange') return;

    moveCurrentParagraphDown(selectedTextRange);
})();


/**
 * カーソルのある段落をひとつ下の段落と入れ替え、カーソル位置を追従させる。
 * @param {TextRange} selectedTextRange - 選択中のテキスト範囲（キャレットのみの場合を含む）
 * @returns {void}
 */
function moveCurrentParagraphDown(selectedTextRange) {
    var story = selectedTextRange.story;
    var paragraphs = story.paragraphs;
    var cursorOffset = selectedTextRange.start;

    var paragraphIndex = getParagraphIndexAtOffset(paragraphs, cursorOffset);

    /* 最終段落は下へ動かせない / The last paragraph cannot move down */
    if (paragraphIndex >= paragraphs.length - 1) return;

    /* 段落内でのカーソル位置を控え、入れ替え後の段落先頭を基準に復元する    */
    /* （改行を数えるかどうかに依存しない） / Independent of how CR is counted */
    var cursorOffsetInParagraph = cursorOffset - paragraphs[paragraphIndex].start;

    swapWithNextParagraph(paragraphs, paragraphIndex);

    restoreCursorPosition(paragraphs[paragraphIndex + 1], cursorOffsetInParagraph);
}


/**
 * 指定した文字オフセットを含む段落のインデックスを返す。
 * @param {Paragraphs} paragraphs - ストーリーの段落コレクション
 * @param {number} cursorOffset - カーソルの文字オフセット
 * @returns {number} 段落インデックス
 */
function getParagraphIndexAtOffset(paragraphs, cursorOffset) {
    for (var i = 0; i < paragraphs.length; i++) {
        if (cursorOffset <= paragraphs[i].end) return i;
    }
    return paragraphs.length - 1;
}


/**
 * 指定した段落を、ひとつ下の段落と入れ替える。
 * カット → 下の段落を複製 → 元の下段落へペースト、の3手で書式ごと交換する。
 * 最終段落を渡してはいけない（カットすると空段落が段落コレクションから消え、
 * 直後の paragraphs[paragraphIndex] が Error 1302 になる）。
 * @param {Paragraphs} paragraphs - ストーリーの段落コレクション
 * @param {number} paragraphIndex - 下へ移動する段落のインデックス（最終段落は不可）
 * @returns {void}
 */
function swapWithNextParagraph(paragraphs, paragraphIndex) {
    /* paragraphs はライブコレクション。cut / duplicate のたびに範囲が変わるため、 */
    /* paragraphs[paragraphIndex] をローカル変数に退避せず毎回引き直す           */
    /* paragraphs is a live collection; re-resolve it after every mutation      */

    /* 対象の段落をクリップボードへ退避し、その位置を空にする */
    paragraphs[paragraphIndex].select();
    app.cut();

    /* 空いた位置に下の段落を複製する */
    paragraphs[paragraphIndex + 1].duplicate(paragraphs[paragraphIndex]);

    /* 元の下段落を、退避しておいた段落で上書きする */
    paragraphs[paragraphIndex + 1].select();
    app.paste();

    /* 再描画しないと直後に読む start が入れ替え前の値のままになることがある */
    /* Without a redraw the offsets read next may still be the pre-swap ones  */
    app.redraw();
}


/**
 * 入れ替え後の段落の中にキャレットを復帰させる。
 * Illustrator にはキャレット位置を直接指定する API が無いため、
 * 目的位置の1文字をカット＆ペーストして選択を作り直す。
 * 改行文字を足場にすると段落が結合してペーストに失敗するので、
 * 段落先頭に向かって最初の通常文字まで戻る。
 * @param {Paragraph} paragraph - 復帰先の段落
 * @param {number} offsetInParagraph - 段落先頭からのカーソルの相対位置
 * @returns {void}
 */
function restoreCursorPosition(paragraph, offsetInParagraph) {
    /* Story に contents は無いので、段落（TextRange）から文字列を取る */
    /* Story has no contents property; read the text from the paragraph range */
    var contents = paragraph.contents;
    if (!contents.length) return;

    var characterOffset = Math.min(offsetInParagraph, contents.length - 1);

    /* 改行を踏まない位置まで段落内で手前へ戻す / Step back to a non-break character */
    while (characterOffset >= 0 && isParagraphBreak(contents.charAt(characterOffset))) {
        characterOffset--;
    }

    /* 空段落は足場になる文字が無い / An empty paragraph has nothing to anchor to */
    if (characterOffset < 0) return;

    paragraph.characters[characterOffset].select();
    app.redraw();
    app.cut();
    app.paste();
    app.redraw();
}


/**
 * 段落区切りとして扱う文字かどうかを判定する。
 * @param {string} character - 判定する1文字
 * @returns {boolean} 段落区切りなら true
 */
function isParagraphBreak(character) {
    var characterCode = character.charCodeAt(0);
    return characterCode == 13 || characterCode == 10 || characterCode == 3;
}
