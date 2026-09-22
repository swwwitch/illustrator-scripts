#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームから、Illustrator標準の「箇条書きと番号付きリスト」を解除します。
テキストの内容・文字属性・段落設定・タブストップは控えて戻すため、リスト書式だけが外れます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearBulletsAndNumbering.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

Removes Illustrator's built-in Bullets and Numbering from the selected text frames.
The text content, character attributes, paragraph settings and tab stops are captured and restored, so only the list formatting comes off.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearBulletsAndNumbering.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClearBulletsAndNumbering";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearBulletsAndNumbering.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearBulletsAndNumbering.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
            noTextFrame: { ja: "テキストフレームを選択してください。", en: "Please select text frames." }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取得する
     * @param {string} labelPath - LABELS のパス（例: "alert.noDocument"）
     * @returns {string} ラベル文字列
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        return labelNode[uiLang] || labelNode.en;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    main();

    /**
     * 前提チェックののち、選択したテキストフレームのリスト書式を解除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length === 0) { alert(getLabel("alert.noSelection")); return; }

        var targetFrames = collectTextFrames(selectedItems, []);
        if (targetFrames.length === 0) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        for (var i = 0; i < targetFrames.length; i++) {
            clearListFormatting(targetFrames[i]);
        }
        app.redraw();
    }

    /**
     * 選択オブジェクトからテキストフレームを集める（グループ内も再帰）
     * @param {PageItem[]} pageItems - 走査対象のオブジェクト配列
     * @param {TextFrame[]} collectedFrames - 収集先の配列
     * @returns {TextFrame[]} 収集したテキストフレーム
     */
    function collectTextFrames(pageItems, collectedFrames) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!pageItem) continue;
            if (pageItem.typename === "TextFrame") {
                collectedFrames.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                /* グループ内のテキストフレームも対象 / include text frames inside groups */
                collectTextFrames(pageItem.pageItems, collectedFrames);
            }
        }
        return collectedFrames;
    }

    /**
     * 1フレームの「箇条書きと番号付きリスト」を解除する
     * contents を入れ直すとリスト書式が外れる（同時に文字書式も初期化されるため、控えてから戻す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function clearListFormatting(textFrame) {
        var frameState = captureFrameState(textFrame);
        if (frameState.contents == null) return;

        try {
            textFrame.contents = frameState.contents;
        } catch (e) {
            return; /* 入れ直せなければ書式も戻さない / leave the frame untouched when the text cannot be reassigned */
        }

        restoreFrameState(textFrame, frameState);
    }

    // =========================================
    // 書式の退避・復元 / Format snapshot & restore
    // =========================================
    // contents の再設定でフレーム全体の書式が初期化されるため、文字属性・段落属性を控えて復元する
    // Setting .contents resets the frame's formatting, so character and paragraph attributes are snapshotted and restored.

    /**
     * テキストと書式の現状を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{contents: (string|null), charSnapshots: Object[], paraSnapshots: Object[]}} 控えた状態
     */
    function captureFrameState(textFrame) {
        var frameState = { contents: null, charSnapshots: [], paraSnapshots: [] };
        try { frameState.contents = textFrame.contents; } catch (e) { return frameState; }
        frameState.charSnapshots = captureCharAttributes(textFrame);
        frameState.paraSnapshots = captureParagraphFormats(textFrame);
        return frameState;
    }

    /**
     * 控えておいた書式をフレームへ復元する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {{charSnapshots: Object[], paraSnapshots: Object[]}} frameState - captureFrameState() が返した控え
     * @returns {void}
     */
    function restoreFrameState(textFrame, frameState) {
        restoreAllCharAttributes(textFrame, frameState.charSnapshots);
        restoreParagraphFormats(textFrame, frameState.paraSnapshots);
    }

    /**
     * フレーム内の全文字の文字属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 文字ごとの属性（文字順）
     */
    function captureCharAttributes(textFrame) {
        var charSnapshots = [];
        try {
            var characters = textFrame.textRange.characters;
            for (var i = 0; i < characters.length; i++) {
                charSnapshots.push(snapshotCharAttributes(characters[i].characterAttributes));
            }
        } catch (e) { }
        return charSnapshots;
    }

    /**
     * 控えた文字属性を全文字へ復元する（文字数は不変なので先頭から順に対応づける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} charSnapshots - captureCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreAllCharAttributes(textFrame, charSnapshots) {
        if (!charSnapshots || charSnapshots.length === 0) return;
        try {
            var characters = textFrame.textRange.characters;
            for (var i = 0; i < characters.length && i < charSnapshots.length; i++) {
                restoreCharAttributes(characters[i].characterAttributes, charSnapshots[i]);
            }
        } catch (e) { }
    }

    /**
     * 1文字分の主要な文字属性を控える
     * @param {CharacterAttributes} characterAttr - 対象の文字属性
     * @returns {Object} 控えた属性
     */
    function snapshotCharAttributes(characterAttr) {
        var charSnapshot = {};
        /* 途中で失敗しても、それまでに読めた属性は残る / attributes read before a failure are kept */
        try {
            charSnapshot.textFont = characterAttr.textFont;
            charSnapshot.size = characterAttr.size;
            charSnapshot.horizontalScale = characterAttr.horizontalScale;
            charSnapshot.verticalScale = characterAttr.verticalScale;
            charSnapshot.baselineShift = characterAttr.baselineShift;
            charSnapshot.tracking = characterAttr.tracking;
            charSnapshot.leading = characterAttr.leading;
            charSnapshot.autoLeading = characterAttr.autoLeading;
            charSnapshot.fillColor = characterAttr.fillColor;
        } catch (e) { }
        return charSnapshot;
    }

    /**
     * 控えた文字属性を1文字へ復元する
     * @param {CharacterAttributes} characterAttr - 復元先の文字属性
     * @param {Object} charSnapshot - snapshotCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreCharAttributes(characterAttr, charSnapshot) {
        if (!charSnapshot) return;
        /* フォントは失敗しやすいので分けて囲み、他の属性の復元を巻き込まない
           Guard the font separately so a failure there does not skip the remaining attributes */
        if (charSnapshot.textFont) { try { characterAttr.textFont = charSnapshot.textFont; } catch (eFont) { } }
        try {
            if (charSnapshot.size != null) characterAttr.size = charSnapshot.size;
            if (charSnapshot.horizontalScale != null) characterAttr.horizontalScale = charSnapshot.horizontalScale;
            if (charSnapshot.verticalScale != null) characterAttr.verticalScale = charSnapshot.verticalScale;
            if (charSnapshot.baselineShift != null) characterAttr.baselineShift = charSnapshot.baselineShift;
            if (charSnapshot.tracking != null) characterAttr.tracking = charSnapshot.tracking;
            /* 行送りは自動行送りより先に戻す（先に autoLeading を立てると固定値が入らない）
               Restore leading before auto-leading (setting auto-leading first would drop the fixed value) */
            if (charSnapshot.leading != null) characterAttr.leading = charSnapshot.leading;
            if (charSnapshot.autoLeading != null) characterAttr.autoLeading = charSnapshot.autoLeading;
            if (charSnapshot.fillColor) characterAttr.fillColor = charSnapshot.fillColor;
        } catch (e) { }
    }

    /**
     * 各段落の段落属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 段落ごとの属性（段落順）
     */
    function captureParagraphFormats(textFrame) {
        var paraSnapshots = [];
        try {
            var paragraphs = textFrame.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) {
                paraSnapshots.push(snapshotParagraphAttributes(paragraphs[i].paragraphAttributes));
            }
        } catch (e) { }
        return paraSnapshots;
    }

    /**
     * 控えた段落属性を各段落へ復元する（段落数は不変なので先頭から順に対応づける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} paraSnapshots - captureParagraphFormats() が返した控え
     * @returns {void}
     */
    function restoreParagraphFormats(textFrame, paraSnapshots) {
        if (!paraSnapshots || paraSnapshots.length === 0) return;
        try {
            var paragraphs = textFrame.paragraphs;
            for (var i = 0; i < paragraphs.length && i < paraSnapshots.length; i++) {
                restoreParagraphAttributes(paragraphs[i].paragraphAttributes, paraSnapshots[i]);
            }
        } catch (e) { }
    }

    /**
     * 1段落分の主要な段落属性を控える
     * @param {ParagraphAttributes} paragraphAttr - 対象の段落属性
     * @returns {Object} 控えた属性
     */
    function snapshotParagraphAttributes(paragraphAttr) {
        var paraSnapshot = {};
        try {
            paraSnapshot.justification = paragraphAttr.justification;
            paraSnapshot.spaceBefore = paragraphAttr.spaceBefore;
            paraSnapshot.spaceAfter = paragraphAttr.spaceAfter;
            paraSnapshot.leftIndent = paragraphAttr.leftIndent;
            paraSnapshot.rightIndent = paragraphAttr.rightIndent;
            paraSnapshot.firstLineIndent = paragraphAttr.firstLineIndent;
            paraSnapshot.tabStops = copyTabStops(paragraphAttr.tabStops);
        } catch (e) { }
        return paraSnapshot;
    }

    /**
     * 控えた段落属性を1段落へ復元する
     * @param {ParagraphAttributes} paragraphAttr - 復元先の段落属性
     * @param {Object} paraSnapshot - snapshotParagraphAttributes() が返した控え
     * @returns {void}
     */
    function restoreParagraphAttributes(paragraphAttr, paraSnapshot) {
        if (!paraSnapshot) return;
        try {
            if (paraSnapshot.justification != null) paragraphAttr.justification = paraSnapshot.justification;
            if (paraSnapshot.spaceBefore != null) paragraphAttr.spaceBefore = paraSnapshot.spaceBefore;
            if (paraSnapshot.spaceAfter != null) paragraphAttr.spaceAfter = paraSnapshot.spaceAfter;
            if (paraSnapshot.leftIndent != null) paragraphAttr.leftIndent = paraSnapshot.leftIndent;
            if (paraSnapshot.rightIndent != null) paragraphAttr.rightIndent = paraSnapshot.rightIndent;
            if (paraSnapshot.firstLineIndent != null) paragraphAttr.firstLineIndent = paraSnapshot.firstLineIndent;
        } catch (e) { }
        /* タブストップは TabStopInfo を作り直して差し替える / rebuild TabStopInfo objects for the tab stops */
        if (paraSnapshot.tabStops) {
            try { paragraphAttr.tabStops = makeTabStops(paraSnapshot.tabStops); } catch (eTab) { }
        }
    }

    /**
     * タブストップを位置と揃えだけの配列として控える
     * @param {TabStopInfo[]} tabStops - 対象のタブストップ
     * @returns {Array<{position: number, alignment: TabStopAlignment}>|null} 控えた内容（取得できなければ null）
     */
    function copyTabStops(tabStops) {
        var tabSpecs = [];
        try {
            for (var i = 0; i < tabStops.length; i++) {
                tabSpecs.push({ position: tabStops[i].position, alignment: tabStops[i].alignment });
            }
        } catch (e) {
            return null;
        }
        return tabSpecs;
    }

    /**
     * 控えた内容から TabStopInfo の配列を作る
     * @param {Array<{position: number, alignment: TabStopAlignment}>} tabSpecs - 控えたタブストップ
     * @returns {TabStopInfo[]} 生成したタブストップ
     */
    function makeTabStops(tabSpecs) {
        var tabStopInfos = [];
        for (var i = 0; i < tabSpecs.length; i++) {
            var tabStopInfo = new TabStopInfo();
            tabStopInfo.alignment = tabSpecs[i].alignment;
            tabStopInfo.position = tabSpecs[i].position;
            tabStopInfos.push(tabStopInfo);
        }
        return tabStopInfos;
    }
})();
