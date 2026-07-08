#targetengine "TypeBasicsPanelEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの基本的な文字組み設定（フォントサイズと行送り・自動カーニング・文字ツメ・
トラッキング・文字揃え）だけをまとめて行う常駐パレットスクリプトです。UnifiedTypePanel.jsx から
これら5機能を抜き出して1画面に集約しています。

- フォントサイズと行送り：サイズ・実質（pt）・行送り（%）の3入力を1パネルに。
  % を段落の「自動行送り量」に設定し常に自動行送りにする（Illustrator 上は常に「自動」表示、
  行送りはサイズに自動追従）。実質＝サイズ×%。実質欄に値を入れると % を逆算して設定。
- 自動カーニング：和文等幅／0／メトリクス／オプティカル（「メトリクス」のみプロポーショナルメトリクスON）
- 文字ツメ：0〜100%。入力欄とスライダー（Shift で 10% 刻み）
- トラッキング：-100〜500。入力欄とスライダー（Shift で 10 刻み）
- 文字揃え：欧文ベースライン／中央／その他（仮想ボディの上下・平均字面の上下をポップアップで選択）
- ラジオや入力を操作すると、その場で選択中のテキストへ即時適用する
- 選択は単体のテキストフレームだけでなく、グループ内のテキストやテキスト編集モードでの範囲選択にも対応
- パレットにフォーカスが戻るたび、または「再読み込み」で選択の現在値を読み取って UI に反映する
- 常駐エンジン（#targetengine）でパレット表示。常駐エンジンの app は
  パレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ
  BridgeTalk で都度委譲する（コードは encodeURIComponent で包んで送信）

### Overview

A docked palette that sets only the basic text-composition attributes (font size & leading,
auto kerning, Tsume, tracking, and character alignment) for the selected text. These five
features are extracted from UnifiedTypePanel.jsx and gathered into a single view.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.2";

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "文字組み基本パネル", en: "Type Basics Panel" }
        },
        field: {
            fontSize: { ja: "サイズ", en: "Size" },
            autoKern: { ja: "自動カーニング", en: "Auto Kerning" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            spacingAdjust: { ja: "字間調整", en: "Letter Spacing" },
            propMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" },
            leading: { ja: "行送り", en: "Leading" },
            leadingPercent: { ja: "行送り%", en: "Leading %" },
            sizeAndLeading: { ja: "フォントサイズと行送り", en: "Font Size & Leading" }
        },
        autoKern: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
        },
        button: {
            reload: { ja: "再読み込み", en: "Reload" },
            auto: { ja: "自動計算", en: "Auto-calc" }
        },
        tip: {
            fontSize: { ja: "選択テキストのフォントサイズ。", en: "Font size of the selection." },
            autoKern: { ja: "自動カーニング方式（和文等幅／0／メトリクス／オプティカル）。", en: "Auto-kerning method (Japanese equal width / 0 / Metrics / Optical)." },
            tsume: { ja: "文字ツメ（0〜100%）。隣接する文字の食い込み量。", en: "Tsume (0–100%): how much adjacent characters tighten." },
            tracking: { ja: "字間（1/1000em）をまとめて調整します（-100〜500）。", en: "Adjusts overall letter spacing in 1/1000 em (-100 to 500)." },
            leading: { ja: "行送り。% を段落の自動行送り量に設定し、行送りはサイズに自動追従します。", en: "Leading: sets the % as the paragraph's auto-leading amount; leading follows the font size." },
            leadingEffective: { ja: "実質の行送り（フォントサイズ×行送り%）。ここに値を入れると % を逆算して設定します。", en: "The effective leading (font size × leading %). Enter a value here to set the leading by back-calculating the %." },
            reload: { ja: "選択中のテキストの現在値を読み取り直して UI に反映します。", en: "Re-read the current values from the selection and reflect them in the UI." },
            leadingAuto: { ja: "現在の行送り（絶対値）からフォントサイズに対する％を計算し、行送り（%）に反映して適用します。", en: "Compute the % of the current (absolute) leading relative to the font size, set it in Leading (%), and apply." },
            proportionalMetrics: { ja: "プロポーショナルメトリクス（和文プロポーショナル字形／メトリクス由来のツメ）をオン／オフします。", en: "Toggle proportional metrics (proportional CJK glyph spacing from font metrics)." }
        },
        msg: {
            applyError: { ja: "適用に失敗しました", en: "Apply failed" }
        }
    };

    /* 言語に応じたラベル文字列を取得 / Resolve a label string for the current language */
    function getLocalizedText(entry) {
        if (!entry) return "";
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(entry) {
        return getLocalizedText(entry) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // 以下の関数群は toString() で連結し、BridgeTalk 本文に同梱して
    // メインエンジン（生きた DOM を持つ）で eval される。常駐エンジン側
    // からは直接呼ばず、本文への埋め込み用途のみ。
    // =========================================

    /* 型名を安全に取得 / Safely resolve a type name */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 例外を短いメッセージ文字列へ / Reduce an exception to a short message string */
    function errMessage(e) {
        return (e && e.message) ? e.message : String(e);
    }

    /* 1 アイテムからテキスト範囲を収集（グループは中を再帰）/ Collect text ranges from one item (descend into groups) */
    function collectTextRangesFromItem(item, selectedRanges) {
        if (!item) return;
        var itemType = getTypeName(item);
        if (itemType === "TextFrame") {
            selectedRanges.push(item.textRange);
        } else if (itemType === "TextRange") {
            selectedRanges.push(item);
        } else if (itemType === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectTextRangesFromItem(item.pageItems[i], selectedRanges);
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            collectTextRangesFromItem(currentSelection[i], selectedRanges);
        }
        return selectedRanges;
    }

    /* 選択が触れている段落を「段落全体の範囲」で取得 / Get the full paragraphs the selection touches */
    function getSelectedParagraphRanges() {
        var ranges = getSelectedTextRanges();
        var paragraphRanges = [];
        for (var i = 0; i < ranges.length; i++) {
            try {
                var paragraphs = ranges[i].paragraphs;
                for (var j = 0; j < paragraphs.length; j++) paragraphRanges.push(paragraphs[j]);
            } catch (e) { }
        }
        return paragraphRanges;
    }

    /* 方式 ID を AutoKernType の列挙値へ変換 / Resolve a method id to an AutoKernType enum value
       和文等幅は欧文のみメトリクス＝和文は等幅（METRICSROMANONLY）*/
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /* 選択範囲にカーニング方式を適用 / Apply a kerning method to the given ranges
       メトリクスのときのみプロポーショナルメトリクスをON、それ以外はOFF */
    function applyKerningToRanges(ranges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.kerningMethod = kerningMethod;
                ranges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    /* 選択範囲にプロポーショナルメトリクスを適用 / Apply proportional metrics to the given ranges */
    function applyPropMetricsToRanges(ranges, useProportionalMetrics) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 選択範囲に文字ツメを適用 / Apply Tsume to the given ranges（0〜100 の百分率）*/
    function applyTsumeToRanges(ranges, tsumePercent) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.Tsume = tsumePercent;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 選択範囲にトラッキングを適用 / Apply tracking to the given ranges（1/1000 em 単位）*/
    function applyTrackingToRanges(ranges, trackingValue) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.tracking = trackingValue;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take this attribute
            }
        }
    }

    /* 自動カーニング method を ID 文字列へ / Convert a kerning method (AutoKernType) to an id string */
    function kernMethodToId(value) {
        var valueText = String(value);
        if (valueText === String(AutoKernType.AUTO)) return "metrics";
        if (valueText === String(AutoKernType.OPTICAL)) return "optical";
        if (valueText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (valueText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /* 文字揃え ID を StyleRunAlignmentType へ変換 / Resolve an alignment id to StyleRunAlignmentType */
    function resolveAlignment(alignId) {
        if (alignId === "top") return StyleRunAlignmentType.top;
        if (alignId === "center") return StyleRunAlignmentType.center;
        if (alignId === "bottom") return StyleRunAlignmentType.bottom;
        if (alignId === "icftop") return StyleRunAlignmentType.icfTop;
        if (alignId === "icfbottom") return StyleRunAlignmentType.icfBottom;
        return StyleRunAlignmentType.ROMANBASELINE;
    }

    /* StyleRunAlignmentType を ID 文字列へ / Convert a StyleRunAlignmentType value to an id string */
    function alignmentToId(value) {
        var valueText = String(value);
        if (valueText === String(StyleRunAlignmentType.top)) return "top";
        if (valueText === String(StyleRunAlignmentType.center)) return "center";
        if (valueText === String(StyleRunAlignmentType.bottom)) return "bottom";
        if (valueText === String(StyleRunAlignmentType.icfTop)) return "icftop";
        if (valueText === String(StyleRunAlignmentType.icfBottom)) return "icfbottom";
        if (valueText === String(StyleRunAlignmentType.ROMANBASELINE)) return "roman";
        return "";
    }

    /* テキスト範囲っぽい型か / Is this a text-range-like type */
    function isTextRangeLikeType(typeName) {
        return typeName === "TextRange" || typeName === "InsertionPoint" || typeName === "Character" ||
            typeName === "Word" || typeName === "Line" || typeName === "Paragraph";
    }

    /* 親をたどって TextFrame を返す / Walk up parents to the enclosing TextFrame */
    function findParentTextFrame(item) {
        var current = item;
        for (var i = 0; i < 20; i++) {
            if (!current) return null;
            if (getTypeName(current) === "TextFrame") return current;
            try { current = current.parent; } catch (e) { return null; }
        }
        return null;
    }

    /* TextFrame の重複判定キー / A dedup key for a TextFrame */
    function getTextFrameKey(frame) {
        try { if (frame.uuid) return frame.uuid; } catch (e) { }
        var bounds = frame.visibleBounds;
        var contents = "";
        try { contents = String(frame.contents); } catch (eContents) { }
        var contentKey = contents.length + ":" + contents.substring(0, 16);
        return frame.typename + ":" + frame.position[0] + ":" + frame.position[1] + ":" + bounds.join(":") + ":" + contentKey;
    }

    /* TextFrame / TextRange から characterAttributes を取得（それ以外は null）/ Get characterAttributes from a TextFrame / TextRange */
    function getAlignmentCharAttrs(item) {
        if (item.typename === "TextFrame") return item.textRange.characterAttributes;
        if (item.typename === "TextRange") return item.characterAttributes;
        return null;
    }

    /* テキストへ再帰的に文字揃えを設定（グループは中を走査）/ Recursively set alignment on text (descends into groups) */
    function applyStyleRunAlignment(item, alignmentValue) {
        if (item.typename === "GroupItem") {
            var applied = 0;
            for (var j = 0; j < item.pageItems.length; j++) applied += applyStyleRunAlignment(item.pageItems[j], alignmentValue);
            return applied;
        }
        var attrs = getAlignmentCharAttrs(item);
        if (attrs === null) return 0;
        attrs.alignment = alignmentValue;
        return 1;
    }

    /* 選択へ文字揃え（縦方向の揃え）を適用 / Apply character alignment to the selection */
    function applyAlignmentToSelection(selection, alignValue) {
        if (getTypeName(selection) === "TextRange") { selection.characterAttributes.alignment = alignValue; return 1; }
        var count = 0;
        if (selection && selection.length) {
            for (var i = 0; i < selection.length; i++) count += applyStyleRunAlignment(selection[i], alignValue);
        }
        return count;
    }

    /* 処理可能なテキストフレームか判定し、該当すれば返す / Return the item if it is a processable TextFrame */
    function getProcessableTextFrame(item) {
        if (!item || item.typename !== "TextFrame") return null;
        if (!item.contents) return null;
        if (!item.lines || item.lines.length === 0) return null;
        return item;
    }

    /* 選択から重複なしの処理対象テキストフレームを収集 / Collect unique processable text frames from the selection */
    function collectLeadingFrames(selectionItems) {
        var frames = [];
        var seenFrameKeys = {};
        var itemList = (selectionItems && selectionItems.typename) ? [selectionItems] : (selectionItems || []);
        for (var i = 0; i < itemList.length; i++) collectLeadingFramesFromItem(itemList[i], frames, seenFrameKeys);
        return frames;
    }

    /* 1 アイテムからフレームを収集（グループは中を走査、テキスト編集モードの範囲は親フレームへ）*/
    function collectLeadingFramesFromItem(item, frames, seenFrameKeys) {
        if (!item) return;
        var typeName = getTypeName(item);
        if (typeName === "TextFrame") { addLeadingFrame(item, frames, seenFrameKeys); return; }
        if (isTextRangeLikeType(typeName)) { addLeadingFrame(findParentTextFrame(item), frames, seenFrameKeys); return; }
        if (typeName === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) collectLeadingFramesFromItem(item.pageItems[i], frames, seenFrameKeys);
        }
    }

    /* 処理可能なフレームを重複なく追加 / Add a processable frame, skipping duplicates */
    function addLeadingFrame(frame, frames, seenFrameKeys) {
        var textFrame = getProcessableTextFrame(frame);
        if (!textFrame) return;
        var frameKey = getTextFrameKey(textFrame);
        if (seenFrameKeys[frameKey]) return;
        seenFrameKeys[frameKey] = true;
        frames.push(textFrame);
    }

    /* 行送りの基準 ID を AutoLeadingType へ変換 / Resolve a leading-basis id to AutoLeadingType */
    function resolveLeadingType(typeId) {
        if (typeId === "baseline") return AutoLeadingType.BOTTOMTOBOTTOM;
        return AutoLeadingType.TOPTOTOP;
    }

    /* AutoLeadingType を行送りの基準 ID へ / Convert an AutoLeadingType to a leading-basis id */
    function leadingTypeToId(typeValue) {
        if (typeValue === AutoLeadingType.BOTTOMTOBOTTOM) return "baseline";
        return "top";
    }

    /* 段落に自動行送りを設定：自動行送り量（%）を段落属性に、各文字を autoLeading=true に
       これにより Illustrator 上は常に「自動」表示となり、行送りはフォントサイズに追従する */
    function applyAutoLeadingToParagraphs(paragraphRanges, percent) {
        for (var i = 0; i < paragraphRanges.length; i++) {
            try {
                paragraphRanges[i].paragraphAttributes.autoLeadingAmount = percent;
                paragraphRanges[i].characterAttributes.autoLeading = true;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    /* 行送りの基準（leadingType）だけを設定し、行送り値は変更しない / Set only the leading basis (leadingType) */
    function applyLeadingBasisToFrames(frames, leadingType) {
        for (var i = 0; i < frames.length; i++) {
            try { frames[i].textRange.leadingType = leadingType; } catch (e) { }
        }
    }

    /* 選択の現在値を読み取り、初期表示用にエンコード / Read the selection's current state for reflection
       戻り値: count|fontSizePt|autoAmount|alignmentId|kernId|tsume|tracking|leadingTypeId|propMetrics|leadingPt
       注意: 値系は「最初のフレームの先頭文字／先頭段落」を代表値として読む簡略化 */
    function readState(frames) {
        var count = frames.length;
        var fontSizePt = NaN, autoAmount = NaN, alignmentId = "", kernId = "", tsume = NaN, tracking = NaN, leadingTypeId = "", propMetrics = 0, leadingPt = NaN;
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var charAttributes = lines[0].characters[0].characterAttributes;
                    fontSizePt = charAttributes.size;
                    leadingPt = charAttributes.leading;
                    alignmentId = alignmentToId(charAttributes.alignment);
                    kernId = kernMethodToId(charAttributes.kerningMethod);
                    tsume = charAttributes.Tsume;
                    tracking = charAttributes.tracking;
                    propMetrics = charAttributes.proportionalMetrics ? 1 : 0;
                    var paragraphs = frames[i].paragraphs;
                    if (paragraphs && paragraphs.length > 0) {
                        autoAmount = paragraphs[0].paragraphAttributes.autoLeadingAmount;
                    }
                    leadingTypeId = leadingTypeToId(frames[i].textRange.leadingType);
                    break;
                }
            } catch (e) { }
        }
        return [count, fontSizePt, autoAmount, alignmentId, kernId, tsume, tracking, leadingTypeId, propMetrics, leadingPt].join("|");
    }

    /* 選択の実際の行送り（絶対値 pt）とフォントサイズを読む / Read the actual leading (absolute pt) and font size
       戻り値: leadingPt|sizePt（先頭フレームの先頭文字を代表値として読む）*/
    function readLeadingAbs(frames) {
        var leadingPt = NaN, sizePt = NaN;
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var charAttributes = lines[0].characters[0].characterAttributes;
                    leadingPt = charAttributes.leading;
                    sizePt = charAttributes.size;
                    break;
                }
            } catch (e) { }
        }
        return leadingPt + "|" + sizePt;
    }

    /* 選択範囲へ作用する定型アクションの共通処理（空チェック→try→redraw→件数返却）*/
    function runRangeAction(ranges, applyFn) {
        if (ranges.length === 0) return "OK:0";
        try { applyFn(); } catch (e) { return "ERR:" + errMessage(e); }
        app.redraw();
        return "OK:" + ranges.length;
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    /* メインエンジンへ送る処理関数（上で定義済みのものを再利用）
       Functions shipped to the main engine (reuse of the helpers above) */
    var WORKER_FUNCS = [
        getTypeName, errMessage, runRangeAction, collectTextRangesFromItem, getSelectedTextRanges, getSelectedParagraphRanges,
        resolveAutoKernType, applyKerningToRanges, applyPropMetricsToRanges, applyTsumeToRanges, applyTrackingToRanges, kernMethodToId,
        resolveAlignment, alignmentToId, getAlignmentCharAttrs, applyStyleRunAlignment, applyAlignmentToSelection,
        isTextRangeLikeType, findParentTextFrame, getTextFrameKey,
        getProcessableTextFrame, collectLeadingFrames, collectLeadingFramesFromItem, addLeadingFrame,
        resolveLeadingType, leadingTypeToId, applyAutoLeadingToParagraphs, applyLeadingBasisToFrames,
        readState, readLeadingAbs
    ];

    /* 関数配列を toString() で連結してソース文字列にする / Join the function array into a source string */
    function buildWorkerLibSource() {
        var src = "";
        for (var i = 0; i < WORKER_FUNCS.length; i++) src += WORKER_FUNCS[i].toString() + "\n";
        return src;
    }
    var workerLibCache = null;
    function getWorkerLibSource() {
        if (workerLibCache === null) workerLibCache = buildWorkerLibSource();
        return workerLibCache;
    }

    /* メインエンジンで実行されるディスパッチャ / Dispatcher executed on the main engine
       結果は "OK:<payload>" または "ERR:<msg>" の文字列で返す */
    function dispatch(action, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        var ranges = getSelectedTextRanges();
        if (action === "count") {
            return "OK:" + ranges.length;
        }
        if (action === "getState") {
            return "OK:" + readState(collectLeadingFrames(app.activeDocument.selection));
        }
        if (action === "getLeadingAbs") {
            return "OK:" + readLeadingAbs(collectLeadingFrames(app.activeDocument.selection));
        }
        // 自動カーニング・文字ツメは段落単位（選択が触れた段落全体へ）/ Auto-kerning & Tsume apply per paragraph
        if (action === "apply") { var kernParas = getSelectedParagraphRanges(); return runRangeAction(kernParas, function () { applyKerningToRanges(kernParas, resolveAutoKernType(params.method)); }); }
        if (action === "applyPropMetrics") { var propParas = getSelectedParagraphRanges(); return runRangeAction(propParas, function () { applyPropMetricsToRanges(propParas, params.propMetrics === 1); }); }
        if (action === "applyTsume") { var tsumeParas = getSelectedParagraphRanges(); return runRangeAction(tsumeParas, function () { applyTsumeToRanges(tsumeParas, params.value); }); }
        // トラッキングはラン（選択）単位のまま / Tracking stays per selection (run)
        if (action === "applyTracking") return runRangeAction(ranges, function () { applyTrackingToRanges(ranges, params.tracking); });
        if (action === "applyFontSize") return runRangeAction(ranges, function () {
            for (var rangeIndex = 0; rangeIndex < ranges.length; rangeIndex++) ranges[rangeIndex].characterAttributes.size = params.sizePt;
        });
        if (action === "applyAlign") {
            var alignSel = app.activeDocument.selection;
            var alignValue = resolveAlignment(params.align);
            var alignCount = 0;
            try {
                alignCount = applyAlignmentToSelection(alignSel, alignValue);
            } catch (errAlign) {
                return "ERR:" + errMessage(errAlign);
            }
            app.redraw();
            return "OK:" + alignCount;
        }
        if (action === "applyLeading") {
            var leadingFrames = collectLeadingFrames(app.activeDocument.selection);
            if (leadingFrames.length === 0) return "OK:0";
            try {
                // 自動行送り量（%）を段落ごとに設定し、常に自動行送りに / Set the auto-leading amount (%) per paragraph; always auto-leading
                applyAutoLeadingToParagraphs(getSelectedParagraphRanges(), params.percent);
                // 行送りの基準（leadingType）はフレーム単位で設定 / The leading basis is set per frame
                if (params.leadingType !== undefined) applyLeadingBasisToFrames(leadingFrames, resolveLeadingType(params.leadingType));
            } catch (errLeading) {
                return "ERR:" + errMessage(errLeading);
            }
            app.redraw();
            return "OK:" + leadingFrames.length;
        }
        if (action === "applyProfile") {
            var profSel = app.activeDocument.selection;
            var profParas = getSelectedParagraphRanges(); // 自動カーニング・文字ツメは段落単位 / Kerning & Tsume per paragraph
            try {
                if (params.kern !== undefined) applyKerningToRanges(profParas, resolveAutoKernType(params.kern));
                if (params.tsume !== undefined) applyTsumeToRanges(profParas, params.tsume);
                if (params.tracking !== undefined) applyTrackingToRanges(ranges, params.tracking);
                if (params.sizePt !== undefined) { for (var sizeRangeIndex = 0; sizeRangeIndex < ranges.length; sizeRangeIndex++) ranges[sizeRangeIndex].characterAttributes.size = params.sizePt; }
                if (params.align !== undefined) applyAlignmentToSelection(profSel, resolveAlignment(params.align));
                if (params.leadingPercent !== undefined) {
                    applyAutoLeadingToParagraphs(getSelectedParagraphRanges(), params.leadingPercent);
                    if (params.leadingType !== undefined) applyLeadingBasisToFrames(collectLeadingFrames(profSel), resolveLeadingType(params.leadingType));
                }
            } catch (errProfile) {
                return "ERR:" + errMessage(errProfile);
            }
            app.redraw();
            return "OK:" + ranges.length;
        }
        return "OK:0";
    }
    var DISPATCH_SRC = "(" + dispatch.toString() + ")";

    /* パラメータを安全な JS リテラル文字列へ / Serialize params to a safe JS literal */
    function paramsToSource(params) {
        if (!params) return "{}";
        var parts = [];
        if (params.method !== undefined) parts.push('method:decodeURIComponent("' + encodeURIComponent(params.method) + '")');
        if (params.value !== undefined) parts.push("value:" + parseInt(params.value, 10));
        if (params.propMetrics !== undefined) parts.push("propMetrics:" + (params.propMetrics ? 1 : 0));
        if (params.tracking !== undefined) parts.push("tracking:" + parseInt(params.tracking, 10));
        if (params.sizePt !== undefined) parts.push("sizePt:" + params.sizePt);
        if (params.align !== undefined) parts.push('align:decodeURIComponent("' + encodeURIComponent(params.align) + '")');
        if (params.kern !== undefined) parts.push('kern:decodeURIComponent("' + encodeURIComponent(params.kern) + '")');
        if (params.tsume !== undefined) parts.push("tsume:" + parseInt(params.tsume, 10));
        if (params.percent !== undefined) parts.push("percent:" + params.percent);
        if (params.leadingPercent !== undefined) parts.push("leadingPercent:" + params.leadingPercent);
        if (params.leadingType !== undefined) parts.push('leadingType:"' + params.leadingType + '"');
        return "{" + parts.join(",") + "}";
    }

    /*
     * 本文をメインエンジンへ送り、結果マーカーを解析して onDone(status, payload) を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープするため、コード全体を
     * encodeURIComponent で包んで送り、ターゲットで decodeURIComponent + eval して復元する。
     */
    function sendWorker(code, onDone) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bridge.onResult = function (response) {
            var payload = response.body || "";
            var colonIndex = payload.indexOf(":");
            var marker = colonIndex >= 0 ? payload.substring(0, colonIndex) : payload;
            var rest = colonIndex >= 0 ? payload.substring(colonIndex + 1) : "";
            if (marker === "OK") onDone("ok", rest);
            else onDone("error", rest);
        };
        bridge.onError = function (response) { onDone("error", response && response.body ? response.body : "BridgeTalk error"); };
        bridge.send();
    }

    /* メインエンジンへアクションを委譲（非同期）/ Delegate an action to the main engine (async) */
    function runWorker(actionId, params, onDone) {
        var code = getWorkerLibSource() + "\nvar __r=" + DISPATCH_SRC + "(\"" + actionId + "\"," + paramsToSource(params) + ");__r;";
        sendWorker(code, onDone);
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [10, 15, 10, 10];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    var UNIT_MAP = { 0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft" };
    var UNIT_TO_PT = {
        0: 72, 1: 2.8346456692913386, 2: 1, 3: 12, 4: 28.346456692913386,
        5: 0.7086614173228346, 6: 1, 7: 72, 8: 2834.6456692913386, 9: 2592, 10: 864
    };

    function getUnitLabel(unitCode) {
        if (unitCode === 5) return "Q";
        return UNIT_MAP[unitCode] || "pt";
    }

    /* 行送りの単位ラベル（Q/H 環境では行送りは「H」）/ Leading unit label */
    function getLeadingUnitLabel(unitCode) {
        if (unitCode === 5) return "H";
        return getUnitLabel(unitCode);
    }

    /* テキスト単位を取得（コード・ラベル・pt 換算係数）/ Resolve the text unit (code, label, pt factor) */
    function getTextUnit() {
        var unitCode = 2;
        try { unitCode = app.preferences.getIntegerPreference("text/units"); } catch (e) { }
        return { code: unitCode, label: getUnitLabel(unitCode), factor: UNIT_TO_PT[unitCode] || 1 };
    }

    /* edittext を ↑↓ キーで増減（Shift=10 / Alt=0.1）/ Step an edittext value with Up/Down keys */
    function changeValueByArrowKey(editText, allowNegative, onUpdate, decimals) {
        function roundToStep(value, step) { return Math.round(value / step) * step; }
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var fieldValue = Number(editText.text);
            if (isNaN(fieldValue)) return;
            var keyboard = ScriptUI.environment.keyboardState;
            var step = 1;
            if (keyboard.shiftKey) step = 10;
            else if (keyboard.altKey) step = 0.1;
            if (keyboard.shiftKey) fieldValue = roundToStep(fieldValue, step);
            if (event.keyName === "Up") fieldValue += step;
            else fieldValue -= step;
            if (keyboard.altKey) fieldValue = Math.round(fieldValue * 10) / 10;
            else fieldValue = Math.round(fieldValue);
            if (!allowNegative && fieldValue < 0) fieldValue = 0;
            event.preventDefault();
            if (typeof decimals === "number" && decimals >= 0) {
                var decimalFactor = Math.pow(10, decimals);
                fieldValue = Math.round(fieldValue * decimalFactor) / decimalFactor;
                editText.text = fieldValue.toFixed(decimals);
            } else {
                editText.text = String(fieldValue);
            }
            if (typeof onUpdate === "function") onUpdate();
        });
    }

    /* 自動カーニングの選択肢を生成 / Build the auto-kerning option list */
    function createAutoKernOptions() {
        return [
            { id: "mono", label: LABELS.autoKern.mono },
            { id: "zero", label: LABELS.autoKern.zero },
            { id: "metrics", label: LABELS.autoKern.metrics },
            { id: "optical", label: LABELS.autoKern.optical }
        ];
    }

    /* 「フォントサイズと行送り」パネルを構築して参照を返す / Build the "Font Size & Leading" panel */
    function buildLeadingPanel(parent, textUnit) {
        var leadingPanel = parent.add("panel", undefined, getLocalizedText(LABELS.field.sizeAndLeading));
        setupPanel(leadingPanel, 6);
        leadingPanel.alignChildren = "left";
        leadingPanel.helpTip = getLocalizedText(LABELS.tip.leading);

        var leadColon = currentLanguage === "ja" ? "：" : ": ";

        // フォントサイズ / Font size
        var leadSizeRow = leadingPanel.add("group");
        leadSizeRow.orientation = "row";
        leadSizeRow.alignChildren = ["left", "center"];
        var leadSizeLabel = leadSizeRow.add("statictext", undefined, getLocalizedText(LABELS.field.fontSize) + leadColon);
        var fontSizeInput = leadSizeRow.add("edittext", undefined, "");
        fontSizeInput.characters = 4;
        leadSizeRow.add("statictext", undefined, textUnit.label);
        leadSizeLabel.helpTip = getLocalizedText(LABELS.tip.fontSize); fontSizeInput.helpTip = leadSizeLabel.helpTip;

        // 実質（フォントサイズ×行送り% の結果。ここに入力すると % を逆算）/ Effective leading (size × %)
        var leadEffectiveRow = leadingPanel.add("group");
        leadEffectiveRow.orientation = "row";
        leadEffectiveRow.alignChildren = ["left", "center"];
        var leadEffectiveLabel = leadEffectiveRow.add("statictext", undefined, getLocalizedText(LABELS.field.leading) + leadColon);
        var leadingEffectiveInput = leadEffectiveRow.add("edittext", undefined, "");
        leadingEffectiveInput.characters = 4;
        leadEffectiveRow.add("statictext", undefined, getLeadingUnitLabel(textUnit.code));
        leadEffectiveLabel.helpTip = getLocalizedText(LABELS.tip.leadingEffective); leadingEffectiveInput.helpTip = leadEffectiveLabel.helpTip;

        // 行送り（自動行送り量 %）/ Leading (auto-leading amount %)
        var leadPercentRow = leadingPanel.add("group");
        leadPercentRow.orientation = "row";
        leadPercentRow.alignChildren = ["left", "center"];
        var leadPercentLabel = leadPercentRow.add("statictext", undefined, getLocalizedText(LABELS.field.leadingPercent) + leadColon);
        var leadingPercentInput = leadPercentRow.add("edittext", undefined, "");
        leadingPercentInput.characters = 4;
        leadPercentRow.add("statictext", undefined, "%");
        var leadingAutoButton = leadPercentRow.add("button", undefined, getLocalizedText(LABELS.button.auto));
        leadingAutoButton.preferredSize.width = 68;
        leadingAutoButton.helpTip = getLocalizedText(LABELS.tip.leadingAuto);
        leadPercentLabel.helpTip = getLocalizedText(LABELS.tip.leading); leadingPercentInput.helpTip = leadPercentLabel.helpTip;

        // ラベル幅を揃える / Unify label widths
        var leadLabelWidth = 68;
        leadSizeLabel.preferredSize.width = leadLabelWidth;
        leadEffectiveLabel.preferredSize.width = leadLabelWidth;
        leadPercentLabel.preferredSize.width = leadLabelWidth;

        return {
            fontSizeInput: fontSizeInput,
            leadingEffectiveInput: leadingEffectiveInput,
            leadingPercentInput: leadingPercentInput,
            leadingAutoButton: leadingAutoButton
        };
    }

    /* 「自動カーニング」パネルを構築して参照を返す / Build the "Auto Kerning" panel */
    function buildAutoKernPanel(parent, autoKernOptions) {
        var autoKernPanel = parent.add("panel", undefined, getLocalizedText(LABELS.field.autoKern));
        setupPanel(autoKernPanel, 6);
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.helpTip = getLocalizedText(LABELS.tip.autoKern);

        var kernRadios = [];
        for (var i = 0; i < autoKernOptions.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLocalizedText(autoKernOptions[i].label));
            kernRadio.value = false;
            kernRadio.index = i;
            kernRadios.push(kernRadio);
        }
        return { kernRadios: kernRadios };
    }

    /* 「字間調整」（文字ツメ・トラッキング）パネルを構築して参照を返す / Build the "Letter Spacing" panel */
    function buildSpacingPanel(parent) {
        var spacingPanel = parent.add("panel", undefined, getLocalizedText(LABELS.field.spacingAdjust));
        setupPanel(spacingPanel, 6);

        // プロポーショナルメトリクス（自動カーニングの「メトリクス」に連動する現在ロジックを可視化）
        // Proportional metrics (surfaces the current logic tied to the "Metrics" kerning option)
        var propMetricsRow = spacingPanel.add("group");
        setupGroup(propMetricsRow, "row");
        propMetricsRow.margins = [0, 0, 0, 4];
        var propMetricsCheckbox = propMetricsRow.add("checkbox", undefined, getLocalizedText(LABELS.field.propMetrics));
        propMetricsCheckbox.helpTip = getLocalizedText(LABELS.tip.proportionalMetrics);

        var tsumeRow = spacingPanel.add("group");
        setupGroup(tsumeRow, "row");
        tsumeRow.add("statictext", undefined, labelText(LABELS.field.tsume));
        var tsumeInput = tsumeRow.add("edittext", undefined, "0");
        tsumeInput.characters = 3;
        tsumeInput.helpTip = getLocalizedText(LABELS.tip.tsume);
        tsumeRow.add("statictext", undefined, "%");
        var tsumeSlider = spacingPanel.add("slider", undefined, 0, 0, 100);

        // 文字ツメとトラッキングの間に少し余白 / A little gap between Tsume and Tracking
        var spacingGap = spacingPanel.add("group");
        spacingGap.preferredSize.height = 1;

        var trackingRow = spacingPanel.add("group");
        setupGroup(trackingRow, "row");
        trackingRow.add("statictext", undefined, labelText(LABELS.field.tracking));
        var trackingInput = trackingRow.add("edittext", undefined, "0");
        trackingInput.characters = 3;
        trackingInput.helpTip = getLocalizedText(LABELS.tip.tracking);
        var trackingSlider = spacingPanel.add("slider", undefined, 0, -100, 500);

        return {
            tsumeInput: tsumeInput,
            tsumeSlider: tsumeSlider,
            trackingInput: trackingInput,
            trackingSlider: trackingSlider,
            propMetricsCheckbox: propMetricsCheckbox
        };
    }

    /* パレットを組み立てて参照を返す（イベント未接続）/ Build the palette and return references (events not wired yet) */
    function createPaletteUI(autoKernOptions) {
        var palette = new Window("palette", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = "fill";

        var textUnit = getTextUnit();

        // 1カラム：フォントサイズと行送り／自動カーニング／字間調整を縦に並べる
        // Single column: font size & leading / auto-kerning / letter spacing, stacked vertically
        var mainColumn = palette.add("group");
        mainColumn.orientation = "column";
        mainColumn.alignChildren = ["fill", "top"];
        mainColumn.spacing = 8;

        var leadingUI = buildLeadingPanel(mainColumn, textUnit);
        var autoKernUI = buildAutoKernPanel(mainColumn, autoKernOptions);
        var spacingUI = buildSpacingPanel(mainColumn);

        // フッター（右＝再読み込み）/ Footer: reload
        var footerGroup = palette.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        var footerSpacer = footerGroup.add("statictext", undefined, "");
        footerSpacer.alignment = ["fill", "center"];
        var reloadButton = footerGroup.add("button", undefined, getLocalizedText(LABELS.button.reload));
        reloadButton.alignment = ["right", "center"];
        reloadButton.helpTip = getLocalizedText(LABELS.tip.reload);

        return {
            palette: palette,
            textUnit: textUnit,
            fontSizeInput: leadingUI.fontSizeInput,
            leadingEffectiveInput: leadingUI.leadingEffectiveInput,
            leadingPercentInput: leadingUI.leadingPercentInput,
            leadingAutoButton: leadingUI.leadingAutoButton,
            kernRadios: autoKernUI.kernRadios,
            tsumeInput: spacingUI.tsumeInput,
            tsumeSlider: spacingUI.tsumeSlider,
            trackingInput: spacingUI.trackingInput,
            trackingSlider: spacingUI.trackingSlider,
            propMetricsCheckbox: spacingUI.propMetricsCheckbox,
            reloadButton: reloadButton
        };
    }

    /* ラジオ配列の排他選択（別コンテナのため手動）/ Exclusive selection across radios in separate containers */
    function selectExclusiveRadio(radios, index) {
        for (var i = 0; i < radios.length; i++) radios[i].value = (i === index);
    }

    /* "count|fontSizePt|autoAmount|alignmentId|kernId|tsume|tracking|leadingTypeId|propMetrics" を分解 / Parse the encoded state string */
    function parseState(rest) {
        var fields = String(rest || "").split("|");
        function toNumber(text) { var parsed = parseFloat(text); return isNaN(parsed) ? NaN : parsed; }
        return {
            count: parseInt(fields[0], 10) || 0,
            fontSizePt: toNumber(fields[1]),
            autoAmount: toNumber(fields[2]),
            alignmentId: fields[3] || "",
            kernId: fields[4] || "",
            tsume: toNumber(fields[5]),
            tracking: toNumber(fields[6]),
            leadingTypeId: fields[7] || "",
            propMetrics: parseInt(fields[8], 10) === 1,
            leadingPt: toNumber(fields[9])
        };
    }

    // =========================================
    // イベント接続 / Wire palette events
    // =========================================
    function bindPaletteEvents(ui, autoKernOptions) {
        // BridgeTalk は非同期なので、連打による委譲の重複を抑止 / Guard against overlapping async delegations
        var workerBusy = false;
        var textUnit = ui.textUnit;

        /* 重要処理の失敗をダイアログで通知 / Surface an important-op failure via an alert */
        function showWorkerError(actionId, payload) {
            var detail = payload ? (": " + String(payload)) : "";
            try { alert("⚠ " + getLocalizedText(LABELS.msg.applyError) + " [" + actionId + "]" + detail); } catch (e) { }
        }

        // 委譲する共通処理（連打抑止）/ Delegate an apply action (guarded against overlap)
        function runApply(actionId, params, onDone) {
            if (workerBusy) return;
            workerBusy = true;
            runWorker(actionId, params, function (status, payload) {
                workerBusy = false;
                if (status === "error") showWorkerError(actionId, payload);
                if (onDone) onDone(status, payload);
            });
        }

        // ---- 自動カーニング / Auto kerning ----
        // id 一致するラジオを選択 / Select the radio whose option id matches
        function selectKernById(targetId) {
            for (var i = 0; i < autoKernOptions.length; i++) {
                if (autoKernOptions[i].id === targetId) { selectExclusiveRadio(ui.kernRadios, i); return; }
            }
        }
        // プロポーショナルメトリクスのチェック状態を反映 / Reflect the proportional-metrics checkbox state
        function reflectPropMetrics(isOn) {
            ui.propMetricsCheckbox.value = !!isOn;
        }
        for (var kernIndex = 0; kernIndex < ui.kernRadios.length; kernIndex++) {
            ui.kernRadios[kernIndex].onClick = function () {
                var methodId = autoKernOptions[this.index].id;
                runApply("apply", { method: methodId });
                // 現在ロジック：メトリクスのときだけプロポーショナルメトリクスON / Current logic: only "metrics" turns it on
                reflectPropMetrics(methodId === "metrics");
            };
        }

        // プロポーショナルメトリクス：チェックで独立して適用（段落単位）/ Proportional metrics: apply independently on toggle
        ui.propMetricsCheckbox.onClick = function () {
            runApply("applyPropMetrics", { propMetrics: this.value });
        };

        // ---- 文字ツメ / Tsume ----
        function reflectTsume(value) {
            ui.tsumeSlider.value = value;
            ui.tsumeInput.text = String(value);
        }
        // 文字ツメスライダー（Shift 併用で 10% 刻み）/ Tsume slider (Shift snaps to 10% steps)
        function tsumeSnap(value) {
            var isShiftPressed = false;
            try { isShiftPressed = ScriptUI.environment.keyboardState.shiftKey; } catch (e) { }
            return isShiftPressed ? Math.round(value / 10) * 10 : Math.round(value);
        }
        ui.tsumeSlider.onChanging = function () {
            var snappedValue = tsumeSnap(this.value);
            this.value = snappedValue;
            ui.tsumeInput.text = String(snappedValue);
        };
        ui.tsumeSlider.onChange = function () {
            var snappedValue = tsumeSnap(this.value);
            this.value = snappedValue;
            ui.tsumeInput.text = String(snappedValue);
            runApply("applyTsume", { value: snappedValue });
        };
        // 文字ツメ入力欄：0〜100 にクランプしてスライダーへ反映 / Tsume input: clamp 0-100 and sync the slider
        function applyTsumeFromInput() {
            var clampedValue = Math.round(parseFloat(ui.tsumeInput.text));
            if (isNaN(clampedValue)) return;
            if (clampedValue < 0) clampedValue = 0; else if (clampedValue > 100) clampedValue = 100;
            ui.tsumeInput.text = String(clampedValue);
            ui.tsumeSlider.value = clampedValue;
            runApply("applyTsume", { value: clampedValue });
        }
        ui.tsumeInput.onChange = applyTsumeFromInput;
        changeValueByArrowKey(ui.tsumeInput, false, applyTsumeFromInput, 0);

        // ---- トラッキング / Tracking ----
        function reflectTracking(value) {
            ui.trackingSlider.value = value;
            ui.trackingInput.text = String(value);
        }
        // トラッキングスライダー（-100〜500、Shift 併用で 10 単位）/ Tracking slider (Shift snaps to 10 steps)
        function trackingSnap(value) {
            var isShiftPressed = false;
            try { isShiftPressed = ScriptUI.environment.keyboardState.shiftKey; } catch (e) { }
            return isShiftPressed ? Math.round(value / 10) * 10 : Math.round(value);
        }
        ui.trackingSlider.onChanging = function () {
            var snappedValue = trackingSnap(this.value);
            this.value = snappedValue;
            ui.trackingInput.text = String(snappedValue);
        };
        ui.trackingSlider.onChange = function () {
            var snappedValue = trackingSnap(this.value);
            this.value = snappedValue;
            ui.trackingInput.text = String(snappedValue);
            runApply("applyTracking", { tracking: snappedValue });
        };
        // トラッキング入力欄：-100〜500 にクランプしてスライダーへ反映 / Tracking input: clamp -100..500 and sync the slider
        function applyTrackingFromInput() {
            var clampedValue = Math.round(parseFloat(ui.trackingInput.text));
            if (isNaN(clampedValue)) return;
            if (clampedValue < -100) clampedValue = -100; else if (clampedValue > 500) clampedValue = 500;
            ui.trackingInput.text = String(clampedValue);
            ui.trackingSlider.value = clampedValue;
            runApply("applyTracking", { tracking: clampedValue });
        }
        ui.trackingInput.onChange = applyTrackingFromInput;
        changeValueByArrowKey(ui.trackingInput, true, applyTrackingFromInput, 0);

        // ---- フォントサイズと行送り / Font size & leading ----
        /* 現在の行送り（自動行送り量 %）を取得 / Read the current leading (auto-leading amount %) */
        function currentLeadingPercent() {
            return parseFloat(ui.leadingPercentInput.text);
        }
        /* 実質行送り（フォントサイズ×%）の表示を更新 / Update the effective-leading display (font size × %) */
        function updateLeadingEffective() {
            var size = parseFloat(ui.fontSizeInput.text);
            var percent = currentLeadingPercent();
            if (isNaN(size) || isNaN(percent)) { ui.leadingEffectiveInput.text = ""; return; }
            ui.leadingEffectiveInput.text = String(Math.round(size * percent / 100 * 10) / 10);
        }
        /* 行送り（%）欄へ値を反映し、実質表示も更新 / Set the leading (%) field and refresh the effective display */
        function reflectLeadingPercent(percent) {
            ui.leadingPercentInput.text = isNaN(percent) ? "" : String(Math.round(percent * 10) / 10);
            updateLeadingEffective();
        }
        /* 行送りを適用（自動行送り量%＋基準は仮想ボディの上で固定）/ Apply leading (auto-leading % with basis fixed to top) */
        function applyLeading() {
            var percent = currentLeadingPercent();
            if (isNaN(percent)) return;
            runApply("applyLeading", { percent: percent, leadingType: "top" });
        }

        // フォントサイズ入力：適用し、実質行送りの表示も更新 / Font size input: apply and refresh the effective leading
        // 行送りは自動行送りなのでサイズ変更に自動追従する（行送りの再適用は不要）
        function applyFontSizeFromInput() {
            var inputValue = parseFloat(ui.fontSizeInput.text);
            if (!isNaN(inputValue)) runApply("applyFontSize", { sizePt: inputValue * textUnit.factor });
            updateLeadingEffective();
        }
        ui.fontSizeInput.onChange = applyFontSizeFromInput;
        ui.fontSizeInput.onChanging = function () { updateLeadingEffective(); };
        changeValueByArrowKey(ui.fontSizeInput, false, applyFontSizeFromInput, 1);

        // 行送り（%）入力：実質表示を更新して適用 / Leading (%) input: refresh the effective display and apply
        function applyLeadingFromPercent() {
            updateLeadingEffective();
            applyLeading();
        }
        ui.leadingPercentInput.onChange = applyLeadingFromPercent;
        ui.leadingPercentInput.onChanging = function () { updateLeadingEffective(); };
        changeValueByArrowKey(ui.leadingPercentInput, false, applyLeadingFromPercent, 1);

        // 実質（pt）入力：フォントサイズから % を逆算して適用 / Effective (pt) input: back-calculate the % from the font size and apply
        function applyLeadingFromEffective() {
            var effective = parseFloat(ui.leadingEffectiveInput.text);
            var size = parseFloat(ui.fontSizeInput.text);
            if (isNaN(effective) || isNaN(size) || size <= 0) return;
            var percent = Math.round((effective / size) * 100 * 10) / 10;
            ui.leadingPercentInput.text = String(percent);
            applyLeading();
        }
        ui.leadingEffectiveInput.onChange = applyLeadingFromEffective;
        changeValueByArrowKey(ui.leadingEffectiveInput, false, applyLeadingFromEffective, 1);

        // 自動計算ボタン：現在の行送り（絶対値）から % を逆算して行送り（%）へ反映し適用
        // Auto-calc button: back-calculate the % from the current (absolute) leading, set Leading (%), and apply
        ui.leadingAutoButton.onClick = function () {
            runWorker("getLeadingAbs", null, function (status, payload) {
                if (status !== "ok") return;
                var parts = String(payload).split("|");
                var leadingPt = parseFloat(parts[0]);
                var sizePt = parseFloat(parts[1]);
                if (isNaN(leadingPt) || isNaN(sizePt) || sizePt <= 0) return;
                var percent = Math.round((leadingPt / sizePt) * 100 * 10) / 10;
                ui.leadingPercentInput.text = String(percent);
                updateLeadingEffective();
                applyLeading();
            });
        };

        /* 選択状態を読み取って UI に反映 / Read the selection state and reflect it into the UI */
        function refreshState() {
            runWorker("getState", null, function (status, payload) {
                if (status !== "ok") return;
                var state = parseState(payload);
                if (state.count <= 0) return;
                // フォントサイズ / Font size
                ui.fontSizeInput.text = isNaN(state.fontSizePt) ? "" : String(Math.round((state.fontSizePt / textUnit.factor) * 10) / 10);
                // 行送り%：自動行送り量% を % 欄に反映（実質欄は現在値で上書きするので計算表示は使わない）
                // Leading %: reflect the auto-leading amount %（the effective field is overwritten by the actual value below）
                ui.leadingPercentInput.text = isNaN(state.autoAmount) ? "" : String(Math.round(state.autoAmount * 10) / 10);
                // 行送り：選択の現在値（絶対値 pt）をそのまま表示（サイズ×% の計算値ではない）
                // Leading: show the selection's actual current value (absolute pt), not the size × % computation
                ui.leadingEffectiveInput.text = isNaN(state.leadingPt) ? "" : String(Math.round((state.leadingPt / textUnit.factor) * 10) / 10);
                // 自動カーニング / Auto kerning
                if (state.kernId) selectKernById(state.kernId);
                // 文字ツメ / Tsume
                if (!isNaN(state.tsume)) reflectTsume(Math.round(state.tsume));
                // トラッキング / Tracking
                if (!isNaN(state.tracking)) reflectTracking(Math.round(state.tracking));
                // プロポーショナルメトリクス / Proportional metrics
                reflectPropMetrics(state.propMetrics);
            });
        }

        // パレット表示・フォーカス復帰のたびに選択状況を反映 / Reflect the selection on show and on regaining focus
        ui.palette.onShow = function () { refreshState(); };
        ui.palette.onActivate = function () { refreshState(); };

        // 再読み込み：選択中のテキストの現在値を読み直して UI に反映 / Reload: re-read the selection's current values into the UI
        ui.reloadButton.onClick = function () { refreshState(); };

        // Esc でパレットを閉じる / Close the palette on Esc
        ui.palette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") ui.palette.close();
        });

        return { refreshState: refreshState };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    // 常駐エンジンに保持するパレット参照のキー / Global key holding the persistent palette instance
    var PALETTE_GLOBAL_KEY = "__typeBasicsPanelInstance";

    function main() {
        // 二重起動防止：既に開いているパレットがあれば新規作成せず前面に出して終了
        // Prevent double launch: if a palette is already open, bring it to the front and exit
        var existingPalette = $.global[PALETTE_GLOBAL_KEY];
        if (existingPalette) {
            try {
                existingPalette.show();
                existingPalette.active = true;
                return;
            } catch (e) {
                // 参照が無効（既に破棄済み）なら作り直す / Stale reference: fall through and rebuild
                $.global[PALETTE_GLOBAL_KEY] = null;
            }
        }

        var autoKernOptions = createAutoKernOptions();
        var ui = createPaletteUI(autoKernOptions);
        bindPaletteEvents(ui, autoKernOptions);

        // パレット参照を常駐エンジンに保持し、閉じたら解放 / Keep the instance on the resident engine; clear it on close
        $.global[PALETTE_GLOBAL_KEY] = ui.palette;
        ui.palette.onClose = function () {
            try { $.global[PALETTE_GLOBAL_KEY] = null; } catch (eClose) { }
        };

        ui.palette.show();
    }

    main();

})();
