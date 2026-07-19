#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの行送り（表示値）が整数で 1 ステップ大きくなるように、自動行送り量（％）を
逆算して設定するスクリプトです。AutoLeadingCalc.jsx を元にした姉妹スクリプトで、ダイアログや
パレットは表示せず、実行するとその場で選択中のテキストへ適用します。
グループ内のテキストやテキスト編集モードの範囲選択にも対応します。

- 段落ごとに、現在の行送りを表示単位（pt / mm / Q(H) など）に換算し、次の整数を目標にする
  （例：26.124 → 27）
- その整数の行送りになるよう「目標行送り ÷ フォントサイズ × 100」で自動行送り量（％）を逆算し、
  autoLeadingAmount に設定して常に自動行送り（autoLeading=true）にする
- 手動行送りの段落も、現在の見た目の行送りを基準に整数へ丸めたうえで自動行送り化される
- 行送りの基準（leadingType）は仮想ボディの上（TOPTOTOP）に固定する
- テキスト編集中に段落内の一部の文字だけを選択している場合は、その段落全体へ適用する
  （行送り量は段落単位の属性のため）

### Overview

Sets the auto-leading amount (%) so the selected text's displayed leading value steps up by one
integer (e.g. 26.124 → 27). A sibling of AutoLeadingCalc.jsx — no dialog is shown; it applies to
the selection in place, including text inside groups and range selections in text-edit mode. For
each paragraph the current leading is converted to the document's text unit, the next integer is
chosen, and the auto-leading amount that yields that integer leading is applied.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoLeadingStep-10";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ステップ数はファイル名末尾の符号付き整数から取得する（例: AutoLeadingStep+10 → 10 / AutoLeadingStep-1 → -1）
// 下記はファイル名から数値を読めなかったときの既定値
// The step is read from the trailing signed integer of the file name (e.g. AutoLeadingStep+10 → 10);
// this is the fallback used only when no number can be read.
var DEFAULT_LEADING_STEP = 1;

(function () {

    var isJa = ($.locale.indexOf("ja") === 0);
    /* 言語に応じた文字列 / Pick a string for the current UI language */
    function t(ja, en) { return isJa ? ja : en; }

    /* テキスト単位コード→pt 換算係数 / Text unit code → points-per-unit factor */
    var UNIT_TO_PT = {
        0: 72, 1: 2.8346456692913386, 2: 1, 3: 12, 4: 28.346456692913386,
        5: 0.7086614173228346, 6: 1, 7: 72, 8: 2834.6456692913386, 9: 2592, 10: 864
    };

    /* ドキュメントのテキスト単位の pt 換算係数を取得 / Get the points-per-unit factor for the text unit */
    function getTextUnitFactor() {
        var unitCode = 2;
        try { unitCode = app.preferences.getIntegerPreference("text/units"); } catch (e) { }
        return UNIT_TO_PT[unitCode] || 1;
    }

    /* 実行中スクリプトのファイル名末尾の符号付き整数をステップ数として取得（読めなければ既定値）
       Read the trailing signed integer of the running script's file name as the step (fallback to default)
       例: AutoLeadingStep+10.jsx → 10 / AutoLeadingStep-1.jsx → -1 / AutoLeadingStep+1.jsx → 1
       @param {number} defaultStep ファイル名から数値を読めないときの既定値
       @returns {number} ステップ数（整数） */
    function getScriptStep(defaultStep) {
        try {
            var baseName = new File($.fileName).name.replace(/\.[^.]*$/, "");
            var matched = baseName.match(/([+\-]?\d+)\s*$/);
            if (matched) {
                var parsed = parseInt(matched[1], 10);
                if (!isNaN(parsed)) return parsed;
            }
        } catch (e) { }
        return defaultStep;
    }

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

    /* 選択から処理対象の TextFrame を収集(グループは再帰)
       Collect processable TextFrames from the selection (recurse into groups) */
    function collectTextFrames(item, frames) {
        if (!item) return;
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") {
            if (item.contents && item.lines && item.lines.length > 0) frames.push(item);
        } else if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectTextFrames(item.pageItems[i], frames);
        }
    }

    /* 範囲が触れている段落(段落全体)を対象配列へ追加 / Add the full paragraphs the range touches */
    function collectParagraphs(range, paragraphTargets) {
        try {
            var paragraphs = range.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) paragraphTargets.push(paragraphs[i]);
        } catch (e) { }
    }

    /* 1段落の行送り(表示値)を整数で step だけ動かす自動行送り量(%)を逆算して適用
       Apply the auto-leading amount (%) that steps the paragraph's displayed leading by `step` integers
       @param {object} paragraph 段落範囲(Paragraph)
       @param {number} unitFactor 表示単位の pt 換算係数
       @param {number} step 行送り(表示単位・整数)のステップ数 */
    function applyLeadingStep(paragraph, unitFactor, step) {
        if (!paragraph.characters || paragraph.characters.length === 0) return;
        try {
            var charAttr = paragraph.characters[0].characterAttributes;
            var sizePt = charAttr.size;
            var leadingPt = charAttr.leading;
            if (isNaN(sizePt) || sizePt <= 0 || isNaN(leadingPt)) return;

            // 現在の行送りを表示単位に換算し、次の整数を目標にする(小さな誤差は吸収)
            // Convert the current leading to display units and target the next integer (absorb tiny float error)
            var currentInUnit = leadingPt / unitFactor;
            var base = (step >= 0) ? Math.floor(currentInUnit + 1e-4) : Math.ceil(currentInUnit - 1e-4);
            var targetInUnit = base + step;
            if (targetInUnit <= 0) return;

            // 目標の整数行送りになる自動行送り量(%)を逆算 / Back-calculate the auto-leading amount (%) for the target integer leading
            var targetLeadingPt = targetInUnit * unitFactor;
            paragraph.paragraphAttributes.autoLeadingAmount = (targetLeadingPt / sizePt) * 100;
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
            // テキスト編集モード：選択が触れている段落だけを対象(一部の文字選択でも段落全体に適用)
            // Text-edit mode: target only the paragraphs the selection touches (partial char selection → whole paragraph)
            collectParagraphs(selection, paragraphTargets);
            var editFrame = findParentTextFrame(selection);
            if (editFrame) typeFrames.push(editFrame);
        } else {
            // 選択ツール：選択したフレーム(グループ内含む)の全段落を対象
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

        // ステップ数はファイル名から取得（例: AutoLeadingStep+10 → 10）/ The step comes from the file name (e.g. +10)
        var step = getScriptStep(DEFAULT_LEADING_STEP);
        // 段落ごとに行送りを整数で step 動かす / Step each paragraph's leading by `step` integers
        var unitFactor = getTextUnitFactor();
        for (var p = 0; p < paragraphTargets.length; p++) applyLeadingStep(paragraphTargets[p], unitFactor, step);
        // 基準は仮想ボディの上に固定(フレーム単位) / Fix the leading basis to the top of the virtual body (per frame)
        for (var g = 0; g < typeFrames.length; g++) {
            try { typeFrames[g].textRange.leadingType = AutoLeadingType.TOPTOTOP; } catch (e) { }
        }
        app.redraw();
    }

    main();

})();
