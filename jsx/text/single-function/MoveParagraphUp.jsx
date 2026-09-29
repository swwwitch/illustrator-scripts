#target illustrator
#targetengine "MoveParagraph"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カーソルのある段落を、ひとつ上の段落と入れ替えます。
sky-chaser-high 氏の moveLineUp.jsx（Visual Studio Code の「行を上へ移動」相当）を、
表示行ではなく段落単位で動かすように改変したものです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MoveParagraphUp.md

### Overview

Swaps the paragraph containing the cursor with the paragraph above it.
A paragraph-based variant of moveLineUp.jsx by sky-chaser-high,
which reproduces Visual Studio Code's "Move Line Up".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MoveParagraphUp.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MoveParagraphUp";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "sky-chaser-high";              /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-08-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MoveParagraphUp.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MoveParagraphUp.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 原作 / Original work
 * @author sky-chaser-high
 * @discussion https://github.com/sky-chaser-high/adobe-illustrator-scripts/blob/main/README_ja.md#%E8%A1%8C%E3%82%92%E4%B8%8A--%E4%B8%8B%E3%81%B8%E7%A7%BB%E5%8B%95
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 動かす単位（1: 段落 / 0: 行。行は自動折り返しの見かけの行も1行に数える）
       / Unit to move (1: paragraph, 0: line; soft-wrapped lines count as separate lines) */
    var MOVE_BY_PARAGRAPH = 1;

    /**
     * 選択中のテキスト範囲を確かめ、カーソルのある段落（または行）を移動する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;

        /* 文字ツールでテキストを選択、またはテキスト内にカーソルがあるときだけ実行する / Run only when text is selected or the caret is inside text */
        var selectedTextRange = app.activeDocument.selection;
        if (!selectedTextRange || selectedTextRange.typename !== "TextRange") return;

        moveCurrentParagraphUp(selectedTextRange);
    }

    /**
     * カーソルのある段落（または行）をひとつ上と入れ替え、カーソル位置を追従させる。
     * @param {TextRange} selectedTextRange - 選択中のテキスト範囲（キャレットのみの場合を含む）
     * @returns {void}
     */
    function moveCurrentParagraphUp(selectedTextRange) {
        var story = selectedTextRange.story;
        var textUnits = MOVE_BY_PARAGRAPH ? story.paragraphs : story.lines;
        var cursorOffset = selectedTextRange.start;

        var unitIndex = getUnitIndexAtOffset(textUnits, cursorOffset);

        /* 先頭は上へ動かせない / The first unit cannot move up */
        if (unitIndex <= 0) return;

        /* 段落（行）内でのカーソル位置を控え、入れ替え後の先頭を基準に復元する  */
        /* （改行を数えるかどうかに依存しない） / Independent of how CR is counted */
        var cursorOffsetInUnit = cursorOffset - textUnits[unitIndex].start;

        /* 「自分を上へ」ではなく「ひとつ上を下へ」動かす。                     */
        /* 最後の段落（行）をカットするとコレクションから消えて落ちるため       */
        /* Move the unit above down instead: cutting the last unit             */
        /* removes it from the collection and breaks the following lookup       */
        swapWithNextUnit(textUnits, unitIndex - 1);

        restoreCursorPosition(textUnits[unitIndex - 1], cursorOffsetInUnit);

        /* 再描画は最後に1回だけ / Redraw once at the end */
        app.redraw();
    }

    /**
     * 指定した文字オフセットを含む段落（行）のインデックスを返す。
     * @param {Paragraphs|Lines} textUnits - ストーリーの段落または行のコレクション
     * @param {number} cursorOffset - カーソルの文字オフセット
     * @returns {number} 段落（行）のインデックス
     */
    function getUnitIndexAtOffset(textUnits, cursorOffset) {
        for (var i = 0; i < textUnits.length; i++) {
            if (cursorOffset <= textUnits[i].end) return i;
        }
        return textUnits.length - 1;
    }

    /**
     * 指定した段落（行）を、ひとつ下と入れ替える。
     * カット → 下を複製 → 元の下へペースト、の3手で書式ごと交換する。
     * 最後の段落（行）を渡してはいけない（カットするとコレクションから消え、
     * 直後の textUnits[unitIndex] が Error 1302 になる）。
     * @param {Paragraphs|Lines} textUnits - ストーリーの段落または行のコレクション
     * @param {number} unitIndex - 下へ移動する段落（行）のインデックス（最後は不可）
     * @returns {void}
     */
    function swapWithNextUnit(textUnits, unitIndex) {
        /* textUnits はライブコレクション。cut / duplicate のたびに範囲が変わるため、 */
        /* textUnits[unitIndex] をローカル変数に退避せず毎回引き直す              */
        /* textUnits is a live collection; re-resolve it after every mutation   */

        /* 対象をクリップボードへ退避し、その位置を空にする */
        textUnits[unitIndex].select();
        app.cut();

        /* 空いた位置に下の段落（行）を複製する */
        textUnits[unitIndex + 1].duplicate(textUnits[unitIndex]);

        /* 元の下を、退避しておいた段落（行）で上書きする */
        textUnits[unitIndex + 1].select();
        app.paste();
    }

    /**
     * 入れ替え後の段落（行）の中にキャレットを復帰させる。
     * Illustrator にはキャレット位置を直接指定する API が無いため、
     * 目的位置の1文字をカット＆ペーストして選択を作り直す。
     * 改行文字を足場にすると段落が結合してペーストに失敗するので、
     * 先頭に向かって最初の通常文字まで戻る。
     * @param {TextRange} textUnit - 復帰先の段落（行）
     * @param {number} offsetInUnit - 段落（行）の先頭からのカーソルの相対位置
     * @returns {void}
     */
    function restoreCursorPosition(textUnit, offsetInUnit) {
        /* Story に contents は無いので、段落（行）の TextRange から文字列を取る */
        /* Story has no contents property; read the text from the unit's range */
        var contents = textUnit.contents;
        if (!contents.length) return;

        var characterOffset = Math.min(offsetInUnit, contents.length - 1);

        /* 改行を踏まない位置まで手前へ戻す / Step back to a non-break character */
        while (characterOffset >= 0 && isParagraphBreak(contents.charAt(characterOffset))) {
            characterOffset--;
        }

        /* 空の段落（行）は足場になる文字が無い / An empty unit has nothing to anchor to */
        if (characterOffset < 0) return;

        textUnit.characters[characterOffset].select();
        app.cut();
        app.paste();
    }

    /**
     * 段落区切りとして扱う文字かどうかを判定する。
     * @param {string} character - 判定する1文字
     * @returns {boolean} 段落区切りなら true
     */
    function isParagraphBreak(character) {
        var characterCode = character.charCodeAt(0);
        return characterCode === 13 || characterCode === 10 || characterCode === 3;
    }

    main();

})();
