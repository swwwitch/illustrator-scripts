#target illustrator
#targetengine "ApplyLeadingPerTextFrame"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択された各テキストフレームの各行について、行頭数文字のフォントサイズを基準に行送りを再計算して適用します。
適用する行送りの割合はダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFramePalette.md

### Overview

Recalculates the leading of each line in the selected text frames from the font size of the first few characters, and applies it.
The leading percentage is chosen in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFramePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyLeadingPerTextFramePalette";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFramePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFramePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var LINE_FONT_SIZE_SAMPLE_COUNT = 5;   /* 行内で参照する文字数 / Characters sampled per line */
    var DEFAULT_AUTO_LEADING_AMOUNT = 175; /* 読み取れないときの自動行送り量（%）/ Auto leading amount (%) when nothing can be read */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PALETTE_MARGINS = 16;              /* パレットの余白 / Palette margins */
    var PALETTE_SPACING = 12;              /* パネル同士の間隔 / Spacing between panels */
    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネルの余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING = 8;                 /* パネル内の間隔 / Spacing inside panels */
    var SHORT_INPUT_CHARACTERS = 3;        /* 行送り・自動行送り量の欄の幅（文字数）/ Width of the leading fields */
    var SPACE_INPUT_CHARACTERS = 4;        /* 段落前後のアキの欄の幅（文字数）/ Width of the paragraph spacing fields */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "行送りの設定", en: "Leading Settings" }
        },
        panel: {
            leading: { ja: "行送り", en: "Leading settings" },
            leadingType: { ja: "行送りの基準", en: "Leading basis" },
            paragraphSpacing: { ja: "段落前後のアキ", en: "Paragraph spacing" }
        },
        radio: {
            topToTop: { ja: "仮想ボディの上基準", en: "Top-to-top (virtual body)" },
            bottomToBottom: { ja: "欧文ベースライン基準", en: "Baseline-to-baseline" },
            otherLeading: { ja: "その他", en: "Other" }
        },
        fieldLabel: {
            spaceBefore: { ja: "段落前", en: "Space before" },
            spaceAfter: { ja: "段落後", en: "Space after" }
        },
        tooltip: {
            autoAmount: {
                ja: "自動行送り量（％）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1",
                en: "Auto leading amount (%). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1"
            },
            leading: {
                ja: "行送り値（［その他］で直接指定）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1",
                en: "Leading value (use Other to set directly). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1"
            },
            space: { ja: "↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1", en: "↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1" },
            otherLeading: {
                ja: "左の行送り値をそのまま使います。段落ごとに、各行の先頭の文字サイズから自動行送り量（％）を逆算して設定します",
                en: "Uses the leading value on the left as is: for each paragraph, the auto leading percentage is worked back from the font size at the start of its lines."
            }
        }
    };

    /**
     * ドットパスでローカライズ文字列を取得する（null 耐性あり）
     * @param {string} labelPath - "panel.leading" のようなドット区切りキー
     * @returns {string} 該当言語の文字列。無ければ英語、さらに無ければ labelPath をそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode == null) return labelPath;
        if (typeof labelNode[uiLang] === "string") return labelNode[uiLang];
        if (typeof labelNode.en === "string") return labelNode.en;
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 行送りの選択肢 / Leading choices
    // =========================================

    var LEADING_CHOICES = [
        { label: "110%", ratio: 1.1, token: "110", isOther: false },
        { label: "125%", ratio: 1.25, token: "125", isOther: false },
        { label: "150%", ratio: 1.5, token: "150", isOther: false },
        { label: getLabel("radio.otherLeading"), ratio: undefined, token: "OTHER", isOther: true }
    ];

    var LEADING_TYPE_CHOICES = [
        { label: getLabel("radio.topToTop"), token: "TOPTOTOP" },
        { label: getLabel("radio.bottomToBottom"), token: "BOTTOMTOBOTTOM" }
    ];

    /**
     * トークンに対応する行送りの選択肢の番号を返す
     * @param {string} choiceToken - "110" / "125" / "150" / "OTHER"
     * @returns {number} 選択肢の番号（見つからなければ 0）
     */
    function getLeadingChoiceIndexByToken(choiceToken) {
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].token === choiceToken) return i;
        }
        return 0;
    }

    /**
     * ［その他］の選択肢の番号を返す
     * @returns {number} 選択肢の番号（無ければ -1）
     */
    function findOtherChoiceIndex() {
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].isOther) return i;
        }
        return -1;
    }

    // =========================================
    // worker 関数 / Worker functions (run in the MAIN engine via BridgeTalk)
    // toString() で送るため JSDoc は付けず、説明は1行の /* */ コメントにする
    // 注意: 内部は // 行コメント禁止・/* */ のみ・必ずセミコロンで終える（toString が改行を消すため）
    // =========================================

    /* 文字の編集中なら、そのテキストフレームを選択し直す / While editing text, select its text frame instead */
    function w_normalizeSelection() {
        if (app.documents.length > 0 && app.selection && app.selection.typename === "TextRange") {
            var story = app.selection.story;
            if (story && story.textFrames.length === 1) {
                var parentTextFrame = story.textFrames[0];
                app.executeMenuCommand("deselectall");
                app.selection = [parentTextFrame];
                try { app.selectTool("Adobe Select Tool"); } catch (e) { }
            }
        }
    }

    /* 文字と行のあるテキストフレームだけを返す（それ以外は null）/ Return the item only when it is a text frame with text and lines */
    function w_getProcessableTextFrame(selectedItem) {
        if (!selectedItem || selectedItem.typename !== "TextFrame") { return null; }
        if (!selectedItem.contents) { return null; }
        if (!selectedItem.lines || selectedItem.lines.length === 0) { return null; }
        return selectedItem;
    }

    /* 選択中の処理できるテキストフレームを重複なく集める / Collect the processable selected text frames without duplicates */
    function w_collectTextFrames() {
        var selectionItems = app.activeDocument.selection;
        var textFrames = [];
        if (!selectionItems || selectionItems.length === 0) { return textFrames; }
        for (var i = 0; i < selectionItems.length; i++) {
            var textFrame = w_getProcessableTextFrame(selectionItems[i]);
            if (!textFrame) { continue; }
            var isDuplicate = false;
            for (var j = 0; j < textFrames.length; j++) { if (textFrames[j] === textFrame) { isDuplicate = true; break; } }
            if (!isDuplicate) { textFrames.push(textFrame); }
        }
        return textFrames;
    }

    /* 行頭から sampleCount 文字ぶんのフォントサイズを集める / Sample the font sizes of the first sampleCount characters of a line */
    function w_sampleLineFontSizes(line, sampleCount) {
        var fontSizes = [];
        if (!line || !line.characters || line.characters.length === 0) { return fontSizes; }
        var maxCount = Math.min(line.characters.length, sampleCount);
        for (var i = 0; i < maxCount; i++) {
            try {
                var fontSize = line.characters[i].characterAttributes.size;
                if (!isNaN(fontSize)) { fontSizes.push(fontSize); }
            } catch (e) { }
        }
        return fontSizes;
    }

    /* 最も多い値を返す（同数なら大きいほう）/ Return the most frequent value (the larger one on a tie) */
    function w_getMostFrequentValue(values) {
        if (!values || values.length === 0) { return NaN; }
        var valueCounts = {};
        var bestValue = values[0];
        var bestCount = 0;
        for (var i = 0; i < values.length; i++) {
            var valueKey = String(values[i]);
            if (!valueCounts[valueKey]) { valueCounts[valueKey] = { value: values[i], count: 0 }; }
            valueCounts[valueKey].count++;
            if (valueCounts[valueKey].count > bestCount) { bestCount = valueCounts[valueKey].count; bestValue = valueCounts[valueKey].value; }
            else if (valueCounts[valueKey].count === bestCount && valueCounts[valueKey].value > bestValue) { bestValue = valueCounts[valueKey].value; }
        }
        return bestValue;
    }

    /* 各行の行頭のフォントサイズを sampledSizes に足していく / Append the line-start font sizes of every line to sampledSizes */
    function w_pushLineSampleSizes(textObject, sampledSizes) {
        var textLines = textObject.lines;
        for (var j = 0; j < textLines.length; j++) {
            var lineSizes = w_sampleLineFontSizes(textLines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
            for (var s = 0; s < lineSizes.length; s++) { sampledSizes.push(lineSizes[s]); }
        }
    }

    /* 段落の基準フォントサイズ（途中で失敗しても集めたぶんで判定）/ Base font size of a paragraph (uses what was sampled even on failure) */
    function w_getParagraphBaseFontSize(paragraph) {
        var sampledSizes = [];
        try {
            w_pushLineSampleSizes(paragraph, sampledSizes);
        } catch (e) { }
        if (sampledSizes.length === 0) { return NaN; }
        return w_getMostFrequentValue(sampledSizes);
    }

    /* テキストフレーム全体の基準フォントサイズ / Base font size of a whole text frame */
    function w_getFrameBaseFontSize(textFrame) {
        var sampledSizes = [];
        w_pushLineSampleSizes(textFrame, sampledSizes);
        if (sampledSizes.length === 0) { return NaN; }
        return w_getMostFrequentValue(sampledSizes);
    }

    /* トークンを AutoLeadingType に変換する / Convert a token to AutoLeadingType */
    function w_resolveLeadingType(leadingTypeToken) {
        if (leadingTypeToken === "BOTTOMTOBOTTOM") { return AutoLeadingType.BOTTOMTOBOTTOM; }
        return AutoLeadingType.TOPTOTOP;
    }

    /* 選択中のテキストフレームに行送りと段落前後のアキを適用し、代表の行送り（pt）を返す / Apply leading and paragraph spacing, then return a representative leading (pt) */
    function w_applyLeading(autoAmount, directMode, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken) {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var textFrames = w_collectTextFrames();
        if (textFrames.length === 0) { return "NOSEL"; }
        var leadingType = w_resolveLeadingType(leadingTypeToken);
        var useDirect = directMode && !isNaN(directLeadingPt);
        var representativePt = NaN;
        try {
            for (var i = 0; i < textFrames.length; i++) {
                var textFrame = textFrames[i];
                textFrame.textRange.characterAttributes.autoLeading = true;
                var frameParagraphs = textFrame.paragraphs;
                for (var p = 0; p < frameParagraphs.length; p++) {
                    try {
                        var paragraph = frameParagraphs[p];
                        var amount;
                        if (useDirect) {
                            var baseFontSize = w_getParagraphBaseFontSize(paragraph);
                            if (isNaN(baseFontSize) || baseFontSize <= 0) { continue; }
                            amount = (directLeadingPt / baseFontSize) * 100;
                        } else {
                            amount = autoAmount;
                        }
                        paragraph.characterAttributes.autoLeading = true;
                        paragraph.paragraphAttributes.spaceBefore = spaceBefore;
                        paragraph.paragraphAttributes.spaceAfter = spaceAfter;
                        paragraph.paragraphAttributes.autoLeadingAmount = amount;
                    } catch (ep) { }
                }
                try { textFrame.textRange.leadingType = leadingType; } catch (et) { }
            }
            if (useDirect) {
                representativePt = directLeadingPt;
            } else {
                var frameBaseFontSize = w_getFrameBaseFontSize(textFrames[0]);
                if (!isNaN(frameBaseFontSize)) { representativePt = frameBaseFontSize * (autoAmount / 100); }
            }
        } catch (e) {
            return "ERR:" + e.message;
        }
        app.redraw();
        return "OK|" + representativePt;
    }

    /* 選択中のテキストフレームから初期値を読む / Read the initial values from the selected text frames */
    function w_readInitial() {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var textFrames = w_collectTextFrames();
        if (textFrames.length === 0) { return "NOSEL"; }
        var autoAmount = DEFAULT_AUTO_LEADING_AMOUNT;
        var leadingPt = NaN;
        var leadingTypeToken = "TOPTOTOP";
        var spaceBefore = 0;
        var spaceAfter = 0;
        var choiceToken = "OTHER";
        var isAuto = false;
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var frameLines = textFrames[i].lines;
                if (frameLines && frameLines.length > 0 && frameLines[0].characters.length > 0) {
                    var firstCharAttrs = frameLines[0].characters[0].characterAttributes;
                    isAuto = firstCharAttrs.autoLeading;
                    if (!isNaN(firstCharAttrs.leading)) { leadingPt = firstCharAttrs.leading; }
                }
            } catch (e) { }
            if (!isNaN(leadingPt)) { break; }
        }
        for (var k = 0; k < textFrames.length; k++) {
            try {
                var frameParagraphs = textFrames[k].paragraphs;
                if (frameParagraphs && frameParagraphs.length > 0) {
                    var paraAttrs = frameParagraphs[0].paragraphAttributes;
                    if (!isNaN(paraAttrs.autoLeadingAmount)) { autoAmount = paraAttrs.autoLeadingAmount; }
                    if (!isNaN(paraAttrs.spaceBefore)) { spaceBefore = paraAttrs.spaceBefore; }
                    if (!isNaN(paraAttrs.spaceAfter)) { spaceAfter = paraAttrs.spaceAfter; }
                    break;
                }
            } catch (e2) { }
        }
        for (var t = 0; t < textFrames.length; t++) {
            try {
                var frameLeadingType = textFrames[t].textRange.leadingType;
                if (frameLeadingType !== undefined && frameLeadingType !== null) {
                    if (frameLeadingType === AutoLeadingType.BOTTOMTOBOTTOM) { leadingTypeToken = "BOTTOMTOBOTTOM"; }
                    else { leadingTypeToken = "TOPTOTOP"; }
                    break;
                }
            } catch (e3) { }
        }
        if (isAuto) {
            var presetTokens = ["110", "125", "150"];
            for (var c = 0; c < presetTokens.length; c++) {
                if (Math.abs(autoAmount - Number(presetTokens[c])) < 0.5) { choiceToken = presetTokens[c]; break; }
            }
        }
        return "OK|" + autoAmount + "|" + leadingPt + "|" + leadingTypeToken + "|" + spaceBefore + "|" + spaceAfter + "|" + choiceToken;
    }

    // worker 関数はすべてここに登録（追加漏れ防止） / Register every worker function here
    var WORKER_FUNCS = [
        w_normalizeSelection,
        w_getProcessableTextFrame,
        w_collectTextFrames,
        w_sampleLineFontSizes,
        w_getMostFrequentValue,
        w_pushLineSampleSizes,
        w_getParagraphBaseFontSize,
        w_getFrameBaseFontSize,
        w_resolveLeadingType,
        w_applyLeading,
        w_readInitial
    ];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================

    var isBusy = false;

    /**
     * worker 関数群のソースを連結して 1 つのコード文字列にする。
     * 先頭で共有定数を宣言し、後続の呼び出し式から参照できるようにする。
     * @returns {string} メインエンジンで eval するソース
     */
    function buildWorkerSource() {
        var workerSource = "var LINE_FONT_SIZE_SAMPLE_COUNT=" + LINE_FONT_SIZE_SAMPLE_COUNT + ";" +
            "var DEFAULT_AUTO_LEADING_AMOUNT=" + DEFAULT_AUTO_LEADING_AMOUNT + ";";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            workerSource += WORKER_FUNCS[i].toString();
        }
        return workerSource;
    }

    /**
     * worker 呼び出し式をメインエンジンへ同期委譲し、マーカー文字列を受け取る。
     * @param {string} callExpr - "w_applyLeading(...)" 等、文字列を返す呼び出し式
     * @returns {string} マーカー（"OK"/"OK|..."/"NODOC"/"NOSEL"/"ERR:..."）
     */
    function runWorker(callExpr) {
        if (isBusy) { return "ERR:busy"; }
        isBusy = true;
        var workerResult = { value: null };
        /* BridgeTalk の送信は失敗しうる / BridgeTalk sending can fail */
        try {
            var workerCode = buildWorkerSource() + "String(" + callExpr + ");";
            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
            bridgeTalk.onResult = function (resultMessage) { workerResult.value = resultMessage.body; };
            bridgeTalk.onError = function (errorMessage) { workerResult.value = "ERR:" + errorMessage.body; };
            bridgeTalk.send(10);
        } catch (e) {
            workerResult.value = "ERR:" + e.message;
        } finally {
            isBusy = false;
        }
        return workerResult.value;
    }

    /**
     * マーカー文字列を解析する。
     * @param {string} resultText - runWorker の戻り値
     * @returns {{ok: boolean, code: string, extra: ?string, msg: string}} 解析結果
     */
    function parseWorkerResult(resultText) {
        if (resultText == null) { return { ok: false, code: "ERR", msg: "no response" }; }
        var head = resultText;
        var extra = null;
        var barIndex = resultText.indexOf("|");
        if (barIndex >= 0) {
            head = resultText.substring(0, barIndex);
            extra = resultText.substring(barIndex + 1);
        }
        if (head === "OK") { return { ok: true, code: "OK", extra: extra }; }
        if (head === "NODOC") { return { ok: false, code: "NODOC" }; }
        if (head === "NOSEL") { return { ok: false, code: "NOSEL" }; }
        if (head.indexOf("ERR") === 0) { return { ok: false, code: "ERR", msg: resultText.substring(4) }; }
        return { ok: false, code: "ERR", msg: resultText };
    }

    // =========================================
    // 数値ユーティリティ / Numeric helpers
    // =========================================

    /**
     * 文字列を数値にする（数値でなければ代わりの値）
     * @param {string} value - 元の文字列
     * @param {number} fallback - 数値でないときの値
     * @returns {number} 数値
     */
    function toNumber(value, fallback) {
        var parsedValue = parseFloat(value);
        return isNaN(parsedValue) ? fallback : parsedValue;
    }

    /**
     * 手入力値をパースし、下限でクランプする（負数の手入力対策）。
     * @param {string} text - 入力欄の文字列
     * @param {number} minValue - 下限
     * @param {number} fallback - パース失敗時の値
     * @returns {number} クランプ後の値
     */
    function clampMinNumber(text, minValue, fallback) {
        var parsedValue = parseFloat(text);
        if (isNaN(parsedValue)) { parsedValue = fallback; }
        if (parsedValue < minValue) { parsedValue = minValue; }
        return parsedValue;
    }

    /**
     * pt の値を文字の単位に換算して小数第1位の文字列にする
     * @param {number} ptValue - pt の値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {string} 換算した文字列（数値でなければ空文字）
     */
    function formatByUnit(ptValue, textUnit) {
        if (isNaN(ptValue) || ptValue === null) { return ""; }
        return (Math.round((ptValue / textUnit.pointsPerUnit) * 10) / 10).toFixed(1);
    }

    // =========================================
    // UI: 矢印キーによる数値増減 / Arrow-key stepping
    // =========================================

    /**
     * ↑↓キーで数値を増減できるようにする（Shift: ±10、Option/Alt: ±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} [onUpdate] - 値を変えたあとに呼ぶ処理
     * @param {number} [decimals] - 表示する小数の桁数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onUpdate, decimals) {
        function roundToStep(value, step) {
            return Math.round(value / step) * step;
        }

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var step = 1;

            if (keyboardState.shiftKey) {
                step = 10;
            } else if (keyboardState.altKey) {
                step = 0.1;
            }

            if (keyboardState.shiftKey) {
                value = roundToStep(value, step);
            }

            if (event.keyName === "Up") {
                value += step;
            } else {
                value -= step;
            }

            if (keyboardState.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();

            if (typeof decimals === "number" && decimals >= 0) {
                var pow = Math.pow(10, decimals);
                value = Math.round(value * pow) / pow;
                editText.text = value.toFixed(decimals);
            } else {
                editText.text = String(value);
            }

            if (typeof onUpdate === "function") onUpdate();
        });
    }

    // =========================================
    // 常駐パレット / Persistent palette
    // =========================================

    /**
     * 選択中のテキストフレームから初期値を読む（読めなければ既定値）
     * @returns {{autoAmount: number, leadingPt: number, leadingTypeToken: string, spaceBefore: number, spaceAfter: number, choiceToken: string}} 初期値
     */
    function readInitialSettings() {
        var initResult = parseWorkerResult(runWorker("w_readInitial()"));
        var initialSettings = {
            autoAmount: DEFAULT_AUTO_LEADING_AMOUNT,
            leadingPt: NaN,
            leadingTypeToken: "TOPTOTOP",
            spaceBefore: 0,
            spaceAfter: 0,
            choiceToken: "110"
        };
        if (initResult.ok && initResult.extra != null) {
            var fields = initResult.extra.split("|");
            initialSettings.autoAmount = toNumber(fields[0], DEFAULT_AUTO_LEADING_AMOUNT);
            initialSettings.leadingPt = toNumber(fields[1], NaN);
            initialSettings.leadingTypeToken = fields[2] || "TOPTOTOP";
            initialSettings.spaceBefore = toNumber(fields[3], 0);
            initialSettings.spaceAfter = toNumber(fields[4], 0);
            initialSettings.choiceToken = fields[5] || "110";
        }
        return initialSettings;
    }

    /**
     * 段落前後のアキの1行（項目名＋数値欄＋単位）を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 項目名のパス
     * @param {number} initialPt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {EditText} 追加した数値欄
     */
    function addSpaceRow(parentPanel, labelPath, initialPt, textUnit) {
        var spaceRow = parentPanel.add("group");
        spaceRow.add("statictext", undefined, labelText(labelPath));
        var spaceInput = spaceRow.add("edittext", undefined, formatByUnit(initialPt, textUnit));
        spaceInput.characters = SPACE_INPUT_CHARACTERS;
        spaceInput.helpTip = getLabel("tooltip.space");
        spaceRow.add("statictext", undefined, textUnit.label);
        return spaceInput;
    }

    /**
     * パレットを組み立てる（イベントはまだ付けない）
     * @param {Object} initialSettings - readInitialSettings() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {Object} パレットと各コントロール
     */
    function buildPalette(initialSettings, textUnit) {
        var leadingPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        leadingPalette.orientation = "column";
        leadingPalette.alignChildren = "fill";
        leadingPalette.margins = PALETTE_MARGINS;
        leadingPalette.spacing = PALETTE_SPACING;

        /* 行送りパネル / Leading panel */
        var leadingPanel = leadingPalette.add("panel", undefined, getLabel("panel.leading"));
        leadingPanel.orientation = "column";
        leadingPanel.alignChildren = "left";
        leadingPanel.margins = PANEL_MARGINS;
        leadingPanel.spacing = PANEL_SPACING;

        var contentGroup = leadingPanel.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["left", "top"];
        contentGroup.spacing = 25;

        /* 左カラム：行送り値（pt 等） / Left column: leading value */
        var leftColumnGroup = contentGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "left";
        leftColumnGroup.spacing = 0;

        var leadingGroup = leftColumnGroup.add("group");
        var leadingInput = leadingGroup.add("edittext", undefined, formatByUnit(initialSettings.leadingPt, textUnit));
        leadingInput.characters = SHORT_INPUT_CHARACTERS;
        leadingInput.helpTip = getLabel("tooltip.leading");
        leadingGroup.add("statictext", undefined, textUnit.label);

        /* 右カラム：自動行送り量（％）＋プリセット / Right column: auto amount (%) and presets */
        var rightColumnGroup = contentGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "left";
        rightColumnGroup.spacing = 6;

        var autoAmountGroup = rightColumnGroup.add("group");
        autoAmountGroup.orientation = "row";
        autoAmountGroup.alignChildren = "center";
        autoAmountGroup.spacing = 6;
        var autoInput = autoAmountGroup.add("edittext", undefined, String(Math.round(initialSettings.autoAmount)));
        autoInput.characters = SHORT_INPUT_CHARACTERS;
        autoInput.helpTip = getLabel("tooltip.autoAmount");
        autoAmountGroup.add("statictext", undefined, "%");

        var leadingChoiceGroup = rightColumnGroup.add("group");
        leadingChoiceGroup.orientation = "column";
        leadingChoiceGroup.alignChildren = "left";
        leadingChoiceGroup.spacing = 6;
        leadingChoiceGroup.margins = [0, 5, 0, 0];

        var leadingRadios = [];
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            leadingRadios.push(leadingChoiceGroup.add("radiobutton", undefined, LEADING_CHOICES[i].label));
            if (LEADING_CHOICES[i].isOther) leadingRadios[i].helpTip = getLabel("tooltip.otherLeading");
        }
        leadingRadios[getLeadingChoiceIndexByToken(initialSettings.choiceToken)].value = true;

        /* 行送りの基準パネル（英語 UI ではタイトルなしのグループ）/ Leading type panel (an untitled group in English) */
        var leadingTypeContainer;
        if (uiLang === "ja") {
            leadingTypeContainer = leadingPalette.add("panel", undefined, getLabel("panel.leadingType"));
            leadingTypeContainer.margins = PANEL_MARGINS;
        } else {
            leadingTypeContainer = leadingPalette.add("group");
            leadingTypeContainer.margins = [0, 0, 0, 0];
        }
        leadingTypeContainer.orientation = "column";
        leadingTypeContainer.alignChildren = "left";
        leadingTypeContainer.spacing = PANEL_SPACING;

        var typeRadios = [];
        var initialTypeIndex = 0;
        for (var t = 0; t < LEADING_TYPE_CHOICES.length; t++) {
            typeRadios.push(leadingTypeContainer.add("radiobutton", undefined, LEADING_TYPE_CHOICES[t].label));
            if (LEADING_TYPE_CHOICES[t].token === initialSettings.leadingTypeToken) { initialTypeIndex = t; }
        }
        typeRadios[initialTypeIndex].value = true;

        /* 段落前後のアキパネル / Paragraph spacing panel */
        var spacePanel = leadingPalette.add("panel", undefined, getLabel("panel.paragraphSpacing"));
        spacePanel.orientation = "column";
        spacePanel.alignChildren = "left";
        spacePanel.margins = PANEL_MARGINS;
        spacePanel.spacing = PANEL_SPACING;

        var spaceBeforeInput = addSpaceRow(spacePanel, "fieldLabel.spaceBefore", initialSettings.spaceBefore, textUnit);
        var spaceAfterInput = addSpaceRow(spacePanel, "fieldLabel.spaceAfter", initialSettings.spaceAfter, textUnit);

        return {
            palette: leadingPalette,
            leadingInput: leadingInput,
            autoInput: autoInput,
            leadingRadios: leadingRadios,
            typeRadios: typeRadios,
            spaceBeforeInput: spaceBeforeInput,
            spaceAfterInput: spaceAfterInput
        };
    }

    /**
     * ラジオボタンの配列から選ばれている番号を返す
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @returns {number} 選ばれている番号（無ければ -1）
     */
    function getSelectedRadioIndex(radioButtons) {
        for (var r = 0; r < radioButtons.length; r++) {
            if (radioButtons[r].value) return r;
        }
        return -1;
    }

    /**
     * UI から適用オプションを読み取る（負数はクランプ）。
     * @param {Object} paletteControls - buildPalette() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {Object} { invalid, directMode, autoAmount, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken }
     */
    function readApplyOptions(paletteControls, textUnit) {
        var choiceIndex = getSelectedRadioIndex(paletteControls.leadingRadios);
        var directMode = (choiceIndex >= 0) && !!LEADING_CHOICES[choiceIndex].isOther;
        var autoAmount = clampMinNumber(paletteControls.autoInput.text, 0, DEFAULT_AUTO_LEADING_AMOUNT);
        var spaceBefore = clampMinNumber(paletteControls.spaceBeforeInput.text, 0, 0) * textUnit.pointsPerUnit;
        var spaceAfter = clampMinNumber(paletteControls.spaceAfterInput.text, 0, 0) * textUnit.pointsPerUnit;
        var typeIndex = getSelectedRadioIndex(paletteControls.typeRadios);
        var leadingTypeToken = LEADING_TYPE_CHOICES[typeIndex >= 0 ? typeIndex : 0].token;
        var directLeadingPt = NaN;
        if (directMode) {
            var leadingValue = parseFloat(paletteControls.leadingInput.text);
            if (isNaN(leadingValue)) { return { invalid: true }; }
            if (leadingValue < 0) { leadingValue = 0; }
            directLeadingPt = leadingValue * textUnit.pointsPerUnit;
        }
        return {
            invalid: false,
            directMode: directMode,
            autoAmount: autoAmount,
            directLeadingPt: directLeadingPt,
            spaceBefore: spaceBefore,
            spaceAfter: spaceAfter,
            leadingTypeToken: leadingTypeToken
        };
    }

    /**
     * 適用オプションから w_applyLeading の呼び出し式を作る
     * @param {Object} applyOptions - readApplyOptions() の戻り値
     * @returns {string} 呼び出し式
     */
    function buildApplyLeadingCall(applyOptions) {
        return "w_applyLeading(" +
            applyOptions.autoAmount + "," +
            applyOptions.directMode + "," +
            applyOptions.directLeadingPt + "," +
            applyOptions.spaceBefore + "," +
            applyOptions.spaceAfter + ",'" +
            applyOptions.leadingTypeToken + "'" +
            ")";
    }

    /**
     * パレットにイベントを付ける
     * @param {Object} paletteControls - buildPalette() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {void}
     */
    function bindPaletteEvents(paletteControls, textUnit) {
        var leadingPalette = paletteControls.palette;
        var leadingInput = paletteControls.leadingInput;
        var autoInput = paletteControls.autoInput;
        var leadingRadios = paletteControls.leadingRadios;
        var isSyncingUI = false;

        /* index 以外の選択を外す（-1 ならすべて外す）/ Select only index (-1 clears all) */
        function selectLeadingChoice(index) {
            for (var r = 0; r < leadingRadios.length; r++) { leadingRadios[r].value = (r === index); }
        }

        /**
         * 現在の UI 値を選択中のテキストフレームへ即適用する。
         * 絶対値で上書きする冪等な処理なので、操作のたびに呼んでも累積しない。
         * @returns {void}
         */
        function applyToSelection() {
            if (isSyncingUI) { return; }
            var applyOptions = readApplyOptions(paletteControls, textUnit);
            if (applyOptions.invalid) { return; }

            var applyResult = parseWorkerResult(runWorker(buildApplyLeadingCall(applyOptions)));
            if (applyResult.ok && !applyOptions.directMode && applyResult.extra != null) {
                var representativePt = parseFloat(applyResult.extra);
                if (!isNaN(representativePt)) {
                    isSyncingUI = true;
                    leadingInput.text = formatByUnit(representativePt, textUnit);
                    isSyncingUI = false;
                }
            }
        }

        /* 自動行送り量を手で変えたらプリセットの選択を外す / Editing the amount clears the preset choice */
        function onAutoAmountEdited() {
            selectLeadingChoice(-1);
            applyToSelection();
        }

        /* 行送り値を手で変えたら［その他］を選ぶ / Editing the leading value selects Other */
        function onLeadingPtEdited() {
            var otherIndex = findOtherChoiceIndex();
            if (otherIndex >= 0) { selectLeadingChoice(otherIndex); }
            applyToSelection();
        }

        for (var rk = 0; rk < leadingRadios.length; rk++) {
            (function (index) {
                leadingRadios[index].onClick = function () {
                    selectLeadingChoice(index);
                    var ratio = LEADING_CHOICES[index].ratio;
                    if (typeof ratio === "number") {
                        isSyncingUI = true;
                        autoInput.text = String(Math.round(ratio * 100));
                        isSyncingUI = false;
                    }
                    applyToSelection();
                };
            })(rk);
        }

        for (var tk = 0; tk < paletteControls.typeRadios.length; tk++) {
            paletteControls.typeRadios[tk].onClick = applyToSelection;
        }

        autoInput.onChange = onAutoAmountEdited;
        leadingInput.onChange = onLeadingPtEdited;
        paletteControls.spaceBeforeInput.onChange = applyToSelection;
        paletteControls.spaceAfterInput.onChange = applyToSelection;

        changeValueByArrowKey(autoInput, false, onAutoAmountEdited);
        changeValueByArrowKey(leadingInput, false, onLeadingPtEdited, 1);
        changeValueByArrowKey(paletteControls.spaceBeforeInput, false, applyToSelection);
        changeValueByArrowKey(paletteControls.spaceAfterInput, false, applyToSelection);

        /* Esc で閉じる / Close on Esc */
        leadingPalette.addEventListener("keydown", function (kbEvent) {
            if (kbEvent.keyName === "Escape") { leadingPalette.close(); }
        });

        leadingPalette.onClose = function () {
            $.global.__ALPTF_PALETTE__ = null;
            return true;
        };
    }

    /**
     * 常駐パレットを表示する（開いているものは閉じてから開き直す）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる（破棄済みだと例外）/ Prevent multiple launches; a disposed palette throws */
        if ($.global.__ALPTF_PALETTE__) {
            try { $.global.__ALPTF_PALETTE__.close(); } catch (e) { }
            $.global.__ALPTF_PALETTE__ = null;
        }

        var textUnit = getUnitInfo("text/units");
        var initialSettings = readInitialSettings();
        var paletteControls = buildPalette(initialSettings, textUnit);
        bindPaletteEvents(paletteControls, textUnit);

        $.global.__ALPTF_PALETTE__ = paletteControls.palette;
        paletteControls.palette.center();
        paletteControls.palette.show();
    }

    showPalette();

})();
