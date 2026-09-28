#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの現在の行送り（絶対値）とフォントサイズから行送り％を逆算し、それを自動行送り量（％）として設定します。
これにより、以後は常に自動行送りとして扱われます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingCalc.md

### Overview

Derives the leading percentage from the current absolute leading and font size of the selected text and stores it as the auto-leading percentage.
From then on the text uses auto leading.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingCalc.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoLeadingCalc";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingCalc.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingCalc.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "テキストが選択されていません。", en: "No text is selected." }
        }
    };

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
            alert(getLabel("alert.noDocument"));
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
            alert(getLabel("alert.noSelection"));
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
