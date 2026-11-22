#target illustrator
#targetengine "MoveParagraph"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.MoveParagraphDown
Meta-Default-language=ja
Name=カーソルのある段落をひとつ下へ移動
Name[en]=Move the paragraph at the cursor down
Author=sky-chaser-high
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-08-27
Tags=text,テキスト
X-com.swwwitch-Script-Name=MoveParagraphDown
X-com.swwwitch-Modified-By=Masahiro Takano (@swwwitch)
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MoveParagraphDown.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MoveParagraphDown.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/MoveParagraphDown.jsx
Description-BEGIN
[ja]
カーソルのある段落を、ひとつ下の段落と入れ替えます（変数 MOVE_BY_PARAGRAPH を 0 にすると行単位）。
sky-chaser-high 氏の moveLineDown.jsx を、段落単位でも動かせるように改変したものです。
[en]
Swaps the paragraph containing the cursor with the paragraph below it (set MOVE_BY_PARAGRAPH to 0 to move lines instead).
A variant of moveLineDown.jsx by sky-chaser-high that can also move whole paragraphs.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

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

        moveCurrentParagraphDown(selectedTextRange);
    }

    /**
     * カーソルのある段落（または行）をひとつ下と入れ替え、カーソル位置を追従させる。
     * @param {TextRange} selectedTextRange - 選択中のテキスト範囲（キャレットのみの場合を含む）
     * @returns {void}
     */
    function moveCurrentParagraphDown(selectedTextRange) {
        var story = selectedTextRange.story;
        var textUnits = MOVE_BY_PARAGRAPH ? story.paragraphs : story.lines;
        var cursorOffset = selectedTextRange.start;

        var unitIndex = getUnitIndexAtOffset(textUnits, cursorOffset);

        /* 最後は下へ動かせない / The last unit cannot move down */
        if (unitIndex >= textUnits.length - 1) return;

        /* 段落（行）内でのカーソル位置を控え、入れ替え後の先頭を基準に復元する  */
        /* （改行を数えるかどうかに依存しない） / Independent of how CR is counted */
        var cursorOffsetInUnit = cursorOffset - textUnits[unitIndex].start;

        swapWithNextUnit(textUnits, unitIndex);

        restoreCursorPosition(textUnits[unitIndex + 1], cursorOffsetInUnit);

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
