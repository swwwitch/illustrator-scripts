#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームから、Illustrator標準の「箇条書きと番号付きリスト」を解除します。
テキストの内容・文字属性・段落設定・タブストップは控えて戻すため、リスト書式だけが外れます。

詳細は README を参照してください。

### Overview

Removes Illustrator's built-in Bullets and Numbering from the selected text frames.
The text content, character attributes, paragraph settings and tab stops are captured and restored, so only the list formatting comes off.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClearBulletsAndNumbering";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-18";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearBulletsAndNumbering.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearBulletsAndNumbering.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語コードを取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        alert: {
            noDoc: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
            noTextFrame: { ja: "テキストフレームを選択してください。", en: "Please select text frames." }
        }
    };

    /**
     * ドット区切りキーで現在の言語のラベルを取得する
     * @param {string} key - LABELS のキー（例: "alert.noDoc"）
     * @returns {string} ラベル文字列
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[currentLanguage] || labelNode["en"];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトからテキストフレームを集める（グループ内も再帰）
     * @param {Array<PageItem>} items - 走査対象のオブジェクト配列
     * @param {Array<TextFrame>} out - 収集先の配列
     * @returns {Array<TextFrame>} 収集したテキストフレーム
     */
    function collectTextFrames(items, out) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            if (item.typename === "TextFrame") {
                out.push(item);
            } else if (item.typename === "GroupItem") {
                // グループ内のテキストフレームも対象 / include text frames inside groups
                collectTextFrames(item.pageItems, out);
            }
        }
        return out;
    }

    main();

    /**
     * 前提チェックののち、選択したテキストフレームのリスト書式を解除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel('alert.noDoc')); return; }
        if (app.selection.length === 0) { alert(getLabel('alert.noSelection')); return; }

        var targetFrames = collectTextFrames(app.selection, []);
        if (targetFrames.length === 0) {
            alert(getLabel('alert.noTextFrame'));
            return;
        }

        for (var i = 0; i < targetFrames.length; i++) {
            clearListFormatting(targetFrames[i]);
        }
        app.redraw();
    }

    /**
     * 1フレームの「箇条書きと番号付きリスト」を解除する
     * contents を入れ直すとリスト書式が外れる（同時に文字書式も初期化されるため、控えてから戻す）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {void}
     */
    function clearListFormatting(frame) {
        var state = captureFrameState(frame);
        if (state.contents == null) return;

        try {
            frame.contents = state.contents;
        } catch (e) {
            return; // 入れ直せなければ書式も戻さない / leave the frame untouched when the text cannot be reassigned
        }

        restoreFrameState(frame, state);
    }

    // =========================================
    // 書式の退避・復元 / Format snapshot & restore
    // =========================================
    // contents の再設定でフレーム全体の書式が初期化されるため、文字属性・段落属性を控えて復元する
    // Setting .contents resets the frame's formatting, so character and paragraph attributes are snapshotted and restored.

    /**
     * テキストと書式の現状を控える
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {{contents: (string|null), charAttrs: Array<object>, paraFormats: Array<object>}} 控えた状態
     */
    function captureFrameState(frame) {
        var state = { contents: null, charAttrs: [], paraFormats: [] };
        try { state.contents = frame.contents; } catch (e) { return state; }
        state.charAttrs = captureCharAttributes(frame);
        state.paraFormats = captureParagraphFormats(frame);
        return state;
    }

    /**
     * 控えておいた書式をフレームへ復元する
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {{charAttrs: Array<object>, paraFormats: Array<object>}} state - captureFrameState() が返した控え
     * @returns {void}
     */
    function restoreFrameState(frame, state) {
        restoreCharAttributesAll(frame, state.charAttrs);
        restoreParagraphFormats(frame, state.paraFormats);
    }

    /**
     * フレーム内の全文字の文字属性を控える
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {Array<object>} 文字ごとの属性（文字順）
     */
    function captureCharAttributes(frame) {
        var attrs = [];
        try {
            var characters = frame.textRange.characters;
            for (var i = 0; i < characters.length; i++) {
                attrs.push(snapshotCharAttributes(characters[i].characterAttributes));
            }
        } catch (e) { }
        return attrs;
    }

    /**
     * 控えた文字属性を全文字へ復元する（文字数は不変なので先頭から順に対応づける）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Array<object>} attrs - captureCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreCharAttributesAll(frame, attrs) {
        if (!attrs || attrs.length === 0) return;
        try {
            var characters = frame.textRange.characters;
            for (var i = 0; i < characters.length && i < attrs.length; i++) {
                restoreCharAttributes(characters[i].characterAttributes, attrs[i]);
            }
        } catch (e) { }
    }

    /**
     * 1文字分の主要な文字属性を控える
     * @param {CharacterAttributes} characterAttr - 対象の文字属性
     * @returns {object} 控えた属性
     */
    function snapshotCharAttributes(characterAttr) {
        var snap = {};
        // 途中で失敗しても、それまでに読めた属性は snap に残る / attributes read before a failure stay in snap
        try {
            snap.textFont = characterAttr.textFont;
            snap.size = characterAttr.size;
            snap.horizontalScale = characterAttr.horizontalScale;
            snap.verticalScale = characterAttr.verticalScale;
            snap.baselineShift = characterAttr.baselineShift;
            snap.tracking = characterAttr.tracking;
            snap.leading = characterAttr.leading;
            snap.autoLeading = characterAttr.autoLeading;
            snap.fillColor = characterAttr.fillColor;
        } catch (e) { }
        return snap;
    }

    /**
     * 控えた文字属性を1文字へ復元する
     * @param {CharacterAttributes} characterAttr - 復元先の文字属性
     * @param {object} snap - snapshotCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreCharAttributes(characterAttr, snap) {
        if (!snap) return;
        // フォントは失敗しやすいので分けて囲み、他の属性の復元を巻き込まないようにする
        // Guard the font separately so a failure there does not skip the remaining attributes
        if (snap.textFont) { try { characterAttr.textFont = snap.textFont; } catch (eFont) { } }
        try {
            if (snap.size != null) characterAttr.size = snap.size;
            if (snap.horizontalScale != null) characterAttr.horizontalScale = snap.horizontalScale;
            if (snap.verticalScale != null) characterAttr.verticalScale = snap.verticalScale;
            if (snap.baselineShift != null) characterAttr.baselineShift = snap.baselineShift;
            if (snap.tracking != null) characterAttr.tracking = snap.tracking;
            // 行送りは自動行送りより先に戻す（先に autoLeading を立てると固定値が入らない）
            // Restore leading before auto-leading (setting auto-leading first would drop the fixed value)
            if (snap.leading != null) characterAttr.leading = snap.leading;
            if (snap.autoLeading != null) characterAttr.autoLeading = snap.autoLeading;
            if (snap.fillColor) characterAttr.fillColor = snap.fillColor;
        } catch (e) { }
    }

    /**
     * 各段落の段落属性を控える
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {Array<object>} 段落ごとの属性（段落順）
     */
    function captureParagraphFormats(frame) {
        var formats = [];
        try {
            var paragraphs = frame.paragraphs;
            for (var p = 0; p < paragraphs.length; p++) {
                formats.push(snapshotParagraphAttributes(paragraphs[p].paragraphAttributes));
            }
        } catch (e) { }
        return formats;
    }

    /**
     * 控えた段落属性を各段落へ復元する（段落数は不変なので先頭から順に対応づける）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Array<object>} formats - captureParagraphFormats() が返した控え
     * @returns {void}
     */
    function restoreParagraphFormats(frame, formats) {
        if (!formats || formats.length === 0) return;
        try {
            var paragraphs = frame.paragraphs;
            for (var p = 0; p < paragraphs.length && p < formats.length; p++) {
                restoreParagraphAttributes(paragraphs[p].paragraphAttributes, formats[p]);
            }
        } catch (e) { }
    }

    /**
     * 1段落分の主要な段落属性を控える
     * @param {ParagraphAttributes} paragraphAttr - 対象の段落属性
     * @returns {object} 控えた属性
     */
    function snapshotParagraphAttributes(paragraphAttr) {
        var snap = {};
        try {
            snap.justification = paragraphAttr.justification;
            snap.spaceBefore = paragraphAttr.spaceBefore;
            snap.spaceAfter = paragraphAttr.spaceAfter;
            snap.leftIndent = paragraphAttr.leftIndent;
            snap.rightIndent = paragraphAttr.rightIndent;
            snap.firstLineIndent = paragraphAttr.firstLineIndent;
            snap.tabStops = copyTabStops(paragraphAttr.tabStops);
        } catch (e) { }
        return snap;
    }

    /**
     * 控えた段落属性を1段落へ復元する
     * @param {ParagraphAttributes} paragraphAttr - 復元先の段落属性
     * @param {object} snap - snapshotParagraphAttributes() が返した控え
     * @returns {void}
     */
    function restoreParagraphAttributes(paragraphAttr, snap) {
        if (!snap) return;
        try {
            if (snap.justification != null) paragraphAttr.justification = snap.justification;
            if (snap.spaceBefore != null) paragraphAttr.spaceBefore = snap.spaceBefore;
            if (snap.spaceAfter != null) paragraphAttr.spaceAfter = snap.spaceAfter;
            if (snap.leftIndent != null) paragraphAttr.leftIndent = snap.leftIndent;
            if (snap.rightIndent != null) paragraphAttr.rightIndent = snap.rightIndent;
            if (snap.firstLineIndent != null) paragraphAttr.firstLineIndent = snap.firstLineIndent;
        } catch (e) { }
        // タブストップは TabStopInfo を作り直して差し替える / rebuild TabStopInfo objects for the tab stops
        if (snap.tabStops) {
            try { paragraphAttr.tabStops = makeTabStops(snap.tabStops); } catch (eTab) { }
        }
    }

    /**
     * タブストップを位置と揃えだけの配列として控える
     * @param {Array<TabStopInfo>} tabStops - 対象のタブストップ
     * @returns {Array<{position: number, alignment: TabStopAlignment}>|null} 控えた内容（取得できなければ null）
     */
    function copyTabStops(tabStops) {
        var copied = [];
        try {
            for (var t = 0; t < tabStops.length; t++) {
                copied.push({ position: tabStops[t].position, alignment: tabStops[t].alignment });
            }
        } catch (e) {
            return null;
        }
        return copied;
    }

    /**
     * 控えた内容から TabStopInfo の配列を作る
     * @param {Array<{position: number, alignment: TabStopAlignment}>} tabSpecs - 控えたタブストップ
     * @returns {Array<TabStopInfo>} 生成したタブストップ
     */
    function makeTabStops(tabSpecs) {
        var tabs = [];
        for (var t = 0; t < tabSpecs.length; t++) {
            var tab = new TabStopInfo();
            tab.alignment = tabSpecs[t].alignment;
            tab.position = tabSpecs[t].position;
            tabs.push(tab);
        }
        return tabs;
    }
})();
