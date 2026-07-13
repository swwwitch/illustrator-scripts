#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの「現在の行送り（絶対値）」とフォントサイズから行送り％を逆算し、
それを自動行送り量（％）として設定して常に自動行送りにするスクリプトです。
TypeBasicsPanel.jsx の「自動計算」ボタンの機能だけを抜き出した単独スクリプトで、
ダイアログやパレットは表示せず、実行するとその場で選択中のテキストへ適用します。
グループ内のテキストやテキスト編集モードの範囲選択にも対応します。

### Overview

Back-calculates the leading percentage from the selection's current (absolute) leading and
font size, then applies it as the paragraph's auto-leading amount. Standalone extract of the
"Auto-calc" button in TypeBasicsPanel.jsx — no dialog is shown; it applies to the selection
in place, including text inside groups and range selections in text-edit mode.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

(function () {

    var isJa = ($.locale.indexOf("ja") === 0);
    /* 言語に応じた文字列 / Pick a string for the current UI language */
    function t(ja, en) { return isJa ? ja : en; }

    /* 型名を安全に取得 / Safely resolve a type name */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try { return obj.constructor ? obj.constructor.name : ""; } catch (e) { return ""; }
    }

    /* 親をたどって TextFrame を返す / Walk up parents to the enclosing TextFrame */
    function findParentTextFrame(item) {
        for (var i = 0; i < 20 && item; i++) {
            if (getTypeName(item) === "TextFrame") return item;
            try { item = item.parent; } catch (e) { return null; }
        }
        return null;
    }

    /* 選択から処理対象の TextFrame を収集（グループは再帰、テキスト編集中の範囲は親フレームへ）
       Collect processable TextFrames from the selection (recurse groups; range → parent frame) */
    function collectTextFrames(item, frames) {
        if (!item) return;
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") {
            if (item.contents && item.lines && item.lines.length > 0) frames.push(item);
        } else if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectTextFrames(item.pageItems[i], frames);
        } else {
            collectTextFrames(findParentTextFrame(item), frames);
        }
    }

    /* 範囲が触れている段落（段落全体）を対象配列へ追加 / Add the full paragraphs the range touches */
    function collectParagraphs(range, paragraphTargets) {
        try {
            var paragraphs = range.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) paragraphTargets.push(paragraphs[i]);
        } catch (e) { }
    }

    /* 1段落へ、その段落の現在の行送り（絶対値）÷サイズから % を逆算して自動行送りを適用
       Apply auto-leading to one paragraph by back-calculating the % from its absolute leading ÷ size */
    function applyAutoLeadingToParagraph(paragraph) {
        if (!paragraph.characters || paragraph.characters.length === 0) return;
        try {
            var charAttr = paragraph.characters[0].characterAttributes;
            var sizePt = charAttr.size;
            var leadingPt = charAttr.leading;
            if (isNaN(sizePt) || sizePt <= 0 || isNaN(leadingPt)) return;
            var percent = Math.round((leadingPt / sizePt) * 100 * 10) / 10;
            paragraph.paragraphAttributes.autoLeadingAmount = percent;
            paragraph.characterAttributes.autoLeading = true;
        } catch (e) { }
    }

    function main() {
        if (app.documents.length === 0) {
            alert(t("ドキュメントを開いてください。", "Please open a document."));
            return;
        }

        var selection = app.activeDocument.selection;
        var paragraphTargets = []; // 対象の段落範囲 / Target paragraph ranges
        var typeFrames = [];       // leadingType を設定するフレーム / Frames to set leadingType on

        if (getTypeName(selection) === "TextRange") {
            // テキスト編集モード：選択が触れている段落だけを対象（一部の文字選択でも段落全体に適用）
            // Text-edit mode: target only the paragraphs the selection touches (partial char selection → whole paragraph)
            collectParagraphs(selection, paragraphTargets);
            var editFrame = findParentTextFrame(selection);
            if (editFrame) typeFrames.push(editFrame);
        } else {
            // 選択ツール：選択したフレーム（グループ内含む）の全段落を対象
            // Selection tool: target every paragraph of the selected frames (including those inside groups)
            var frames = [];
            var items = selection || [];
            for (var i = 0; i < items.length; i++) collectTextFrames(items[i], frames);
            for (var f = 0; f < frames.length; f++) {
                collectParagraphs(frames[f].textRange, paragraphTargets);
                typeFrames.push(frames[f]);
            }
        }

        if (paragraphTargets.length === 0) {
            alert(t("テキストが選択されていません。", "No text is selected."));
            return;
        }

        // 段落ごとに現在の行送りから % を逆算して自動行送りを適用 / Apply auto-leading per paragraph
        for (var t2 = 0; t2 < paragraphTargets.length; t2++) applyAutoLeadingToParagraph(paragraphTargets[t2]);
        // 基準は仮想ボディの上に固定（フレーム単位）/ Fix the leading basis to the top of the virtual body (per frame)
        for (var g = 0; g < typeFrames.length; g++) {
            try { typeFrames[g].textRange.leadingType = AutoLeadingType.TOPTOTOP; } catch (e) { }
        }
        app.redraw();
    }

    main();

})();
