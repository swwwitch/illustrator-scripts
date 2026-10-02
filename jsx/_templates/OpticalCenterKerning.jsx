#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

中央揃えのポイント文字で、行ごとの字面の中心を揃えの基準に寄せる再利用テンプレートです。
行頭のカーニングで補正し、行末の「、」などの空きで寄って見える行を、補正の強さ（％）の分だけ戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OpticalCenterKerning.md

### Overview

A reusable template that pulls each line's glyph center of centered point text toward the alignment point.
It compensates with kerning at the start of the line, scaled by a correction strength (%).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OpticalCenterKerning.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "OpticalCenterKerning";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-03";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OpticalCenterKerning.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OpticalCenterKerning.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var CORRECTION_STRENGTH = 35;  /* 補正の強さ（%）。100で字面の中心をぴったり揃える / correction strength (%); 100 centers the glyphs exactly */
    var MIN_KERNING         = 1;   /* これ未満の補正は入れない（1/1000 em） / skip corrections smaller than this (1/1000 em) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noText: { ja: "中央揃えのポイント文字が選択されていません。", en: "No centered point text is selected." }
        }
    };

    // =========================================
    // 字面の中央揃え / Optical center kerning
    // =========================================

    // 字面の中央揃えの使い方 / How to reuse the optical center kerning part
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る。
    //    識別子は OPTICAL_CENTER_KERNING_DEFAULTS / getTextLineRanges / measureLineInk / calcOpticalCenterKerning /
    //    getOpticalCenterKernings / applyOpticalCenterKerning
    // 2. 書き込みまで: applyOpticalCenterKerning(textFrame, { strength: 35, minKerning: 1 })。補正した行数を返す
    //    値だけ欲しいとき（表示・プレビュー用）: getOpticalCenterKernings(textFrame, options) が
    //    [{ lineStart: 行頭の文字番号, kerning: 値 }, …] を返す（テキストは変えない）
    // 3. 対象は中央揃えのポイント文字の行だけ。エリア内文字・パス上文字は 0 行（[]）で返る
    // 4. 測定のたびに複製を作ってアウトライン化する（すぐ消す）。選択は呼ぶ前に控えておく

    // 字面の中央揃え（再利用パーツ） / Optical center kerning (reusable)

    /* 既定値。options で上書きできる / Defaults; override them with options */
    var OPTICAL_CENTER_KERNING_DEFAULTS = {
        strength: 35,   /* 補正の強さ（%）。100で字面の中心をぴったり揃える / correction strength (%); 100 centers the glyphs exactly */
        minKerning: 1   /* これ未満の補正は0にする（1/1000 em） / corrections smaller than this become 0 (1/1000 em) */
    };

    /**
     * 本文を改行で区切り、行ごとの文字番号の範囲を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {Array} [{ start: 行頭の文字番号, end: 改行の位置（含まない） }, …]
     */
    function getTextLineRanges(textFrame) {
        var lineTexts = textFrame.contents.split("\r");
        var ranges = [];
        var start = 0;
        for (var i = 0; i < lineTexts.length; i++) {
            ranges.push({ start: start, end: start + lineTexts[i].length });
            start += lineTexts[i].length + 1;
        }
        return ranges;
    }

    /**
     * 1行分だけを残した複製をアウトライン化し、字面の左右を測る
     * @param {TextFrame} textFrame - 対象のポイント文字
     * @param {number} lineStart - 行頭の文字番号
     * @param {number} lineEnd - 行末の次（改行の位置）の文字番号
     * @returns {Array|null} [左, 右]。字形が無ければ null
     */
    function measureLineInk(textFrame, lineStart, lineEnd) {
        var dup = textFrame.duplicate();

        /* 対象行以外の文字を右から消す / Remove the other lines' characters from the right */
        for (var i = dup.characters.length - 1; i >= lineEnd; i--) {
            dup.characters[i].remove();
        }
        for (var j = lineStart - 1; j >= 0; j--) {
            dup.characters[j].remove();
        }

        /* 補正前の状態で測る / Measure without the existing correction */
        try { dup.characters[0].kerning = 0; } catch (e) {}

        /* createOutline は複製を消費する / createOutline consumes the duplicate */
        var outline = dup.createOutline();
        var bounds = outline.geometricBounds;
        var hasGlyphs = outline.pageItems.length > 0;
        outline.remove();

        return hasGlyphs ? [bounds[0], bounds[2]] : null;
    }

    /**
     * 字面の中心を揃えの基準に寄せるカーニング値（1/1000 em）を求める
     * 中央揃えで行頭に k を入れると行幅が k 伸び、字面は k/2 右へ動く
     * @param {number} offset - 基準X − 字面の中心（pt）
     * @param {number} fontSize - 行頭の文字サイズ（pt）
     * @param {number} strength - 補正の強さ（%）
     * @returns {number} 強さを掛けて丸めたカーニング値
     */
    function calcOpticalCenterKerning(offset, fontSize, strength) {
        return Math.round(2 * offset / fontSize * 1000 * strength / 100);
    }

    /**
     * 中央揃えの行ごとに、行頭へ入れるカーニング値を求める（テキストは変えない）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} [options] - { strength, minKerning }。省いた値は OPTICAL_CENTER_KERNING_DEFAULTS
     * @returns {Array} [{ lineStart: 行頭の文字番号, kerning: 値 }, …]。対象外のテキストは []
     */
    function getOpticalCenterKernings(textFrame, options) {
        var results = [];
        if (textFrame.kind !== TextType.POINTTEXT) return results;

        var opts = options || {};
        var strength = (opts.strength != null) ? opts.strength : OPTICAL_CENTER_KERNING_DEFAULTS.strength;
        var minKerning = (opts.minKerning != null) ? opts.minKerning : OPTICAL_CENTER_KERNING_DEFAULTS.minKerning;

        var characters = textFrame.characters;
        var centerX = textFrame.anchor[0];
        var ranges = getTextLineRanges(textFrame);

        for (var i = 0; i < ranges.length; i++) {
            var range = ranges[i];

            /* 空行と中央揃え以外の行は飛ばす / Skip empty lines and lines that are not centered */
            if (range.end <= range.start) continue;
            if (characters[range.start].paragraphAttributes.justification !== Justification.CENTER) continue;

            var ink = measureLineInk(textFrame, range.start, range.end);
            if (!ink) continue;

            var fontSize = characters[range.start].characterAttributes.size;
            var kerning = calcOpticalCenterKerning(centerX - (ink[0] + ink[1]) / 2, fontSize, strength);
            results.push({ lineStart: range.start, kerning: (Math.abs(kerning) < minKerning) ? 0 : kerning });
        }
        return results;
    }

    /**
     * 中央揃えの行ごとに、行頭のカーニングで字面の中心を補正する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} [options] - getOpticalCenterKernings と同じ
     * @returns {number} 補正した行数
     */
    function applyOpticalCenterKerning(textFrame, options) {
        /* 先にすべての行を測ってから書き込む / Measure every line before writing */
        var kernings = getOpticalCenterKernings(textFrame, options);
        for (var i = 0; i < kernings.length; i++) {
            textFrame.characters[kernings[i].lineStart].kerning = kernings[i].kerning;
        }
        return kernings.length;
    }

    // 字面の中央揃え（再利用パーツ）ここまで / End of the reusable optical center kerning

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提チェックののち、選択したポイント文字を補正する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 処理中に複製を作るので選択の控えを取る / Snapshot the selection since duplicates are created */
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }

        var options = { strength: CORRECTION_STRENGTH, minKerning: MIN_KERNING };
        var adjustedCount = 0;
        for (var j = 0; j < selectedItems.length; j++) {
            var item = selectedItems[j];
            if (item.typename !== "TextFrame" || item.locked || item.hidden) continue;
            if (applyOpticalCenterKerning(item, options) > 0) adjustedCount++;
        }

        if (adjustedCount === 0) {
            alert(getLabel("alert.noText"));
        }
    }

    main();

})();
