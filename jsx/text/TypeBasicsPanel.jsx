#targetengine "TypeBasicsPanelEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの基本的な文字組み設定（フォントサイズと行送り・自動カーニング・プロポーショナルメトリクス・文字ツメ・トラッキング）だけをまとめて行う常駐パレットです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypeBasicsPanel.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n29e7115b5e70

### Overview

A persistent palette covering just the basics of typography for the selected text: font size and leading, auto-kerning, proportional metrics, tsume and tracking.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeBasicsPanel.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypeBasicsPanel";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypeBasicsPanel.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeBasicsPanel.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n29e7115b5e70"; /* 紹介記事 / article URL */
var SCRIPT_PRO_URL = "https://note.com/dtp_tranist/n/n4e2b79cf2891"; /* 上位版 / advanced version */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS         = [10, 15, 10, 10];  /* パネルの余白 / panel margins */
    var PANEL_SPACING         = 8;                 /* パネル・行の既定の間隔 / default panel and row spacing */
    var INNER_PANEL_SPACING   = 6;                 /* 各パネル内の間隔 / spacing inside each panel */
    var MAIN_COLUMN_SPACING   = 8;                 /* パネルどうしの間隔 / spacing between panels */
    var LEADING_LABEL_WIDTH   = 68;                /* サイズ／行送りの項目名の幅 / width of the size and leading labels */
    var AUTO_BUTTON_WIDTH     = 68;                /* ［自動計算］ボタンの幅 / width of the Auto-calc button */
    var LEADING_FIELD_CHARS   = 4;                 /* サイズ／行送りの入力欄の幅 / width of the size and leading fields */
    var SPACING_FIELD_CHARS   = 3;                 /* 文字ツメ／トラッキングの入力欄の幅 / width of the tsume and tracking fields */
    var PROP_METRICS_MARGINS  = [0, 0, 0, 4];      /* プロポーショナルメトリクス行の余白 / margins of the proportional metrics row */

    // =========================================
    // 値の範囲 / Value ranges
    // =========================================
    var TSUME_MIN     = 0;    /* 文字ツメ（%） / tsume (%) */
    var TSUME_MAX     = 100;
    var TRACKING_MIN  = -100; /* トラッキング（1/1000 em） / tracking (1/1000 em) */
    var TRACKING_MAX  = 500;

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

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "文字組み基本パネル", en: "Type Basics Panel" }
        },
        panel: {
            sizeAndLeading: { ja: "フォントサイズと行送り", en: "Font Size & Leading" },
            autoKern: { ja: "自動カーニング", en: "Auto Kerning" },
            spacingAdjust: { ja: "字間調整", en: "Letter Spacing" }
        },
        fieldLabel: {
            fontSize: { ja: "サイズ", en: "Size" },
            leading: { ja: "行送り", en: "Leading" },
            leadingPercent: { ja: "行送り%", en: "Leading %" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            tracking: { ja: "トラッキング", en: "Tracking" }
        },
        checkbox: {
            propMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" }
        },
        radio: {
            kernMono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            kernZero: { ja: "0", en: "0" },
            kernMetrics: { ja: "メトリクス", en: "Metrics" },
            kernOptical: { ja: "オプティカル", en: "Optical" }
        },
        button: {
            reload: { ja: "再読み込み", en: "Reload" },
            autoCalc: { ja: "自動計算", en: "Auto-calc" }
        },
        tooltip: {
            fontSize: { ja: "選択テキストのフォントサイズ。", en: "Font size of the selection." },
            autoKern: {
                ja: "自動カーニング方式（和文等幅／0／メトリクス／オプティカル）。",
                en: "Auto-kerning method (Japanese equal width / 0 / Metrics / Optical)."
            },
            kernMono: {
                ja: "欧文だけメトリクスで詰め、和文は等幅のままにします。プロポーショナルメトリクスはオフにします。",
                en: "Kerns only Roman text by its metrics and keeps Japanese equal width. Turns proportional metrics off."
            },
            kernZero: {
                ja: "自動カーニングを使いません。プロポーショナルメトリクスはオフにします。",
                en: "Turns auto-kerning off. Also turns proportional metrics off."
            },
            kernMetrics: {
                ja: "フォントのメトリクスで詰め、プロポーショナルメトリクスもオンにします。",
                en: "Kerns by the font's metrics and turns proportional metrics on."
            },
            kernOptical: {
                ja: "字形の形から自動で詰めます。プロポーショナルメトリクスはオフにします。",
                en: "Kerns from the glyph shapes. Turns proportional metrics off."
            },
            tsume: { ja: "文字ツメ（0〜100%）。隣接する文字の食い込み量。", en: "Tsume (0–100%): how much adjacent characters tighten." },
            tsumeSlider: {
                ja: "文字ツメを 0〜100% で調整します。Shift キーを押しながら動かすと 10% 刻みになります。",
                en: "Adjusts tsume from 0 to 100%. Hold Shift while dragging to snap to steps of 10%."
            },
            tracking: { ja: "字間（1/1000em）をまとめて調整します（-100〜500）。", en: "Adjusts overall letter spacing in 1/1000 em (-100 to 500)." },
            trackingSlider: {
                ja: "トラッキングを -100〜500 で調整します。Shift キーを押しながら動かすと 10 刻みになります。",
                en: "Adjusts tracking from -100 to 500. Hold Shift while dragging to snap to steps of 10."
            },
            leading: {
                ja: "行送り。% を段落の自動行送り量に設定し、行送りはサイズに自動追従します。",
                en: "Leading: sets the % as the paragraph's auto-leading amount; leading follows the font size."
            },
            leadingEffective: {
                ja: "実質の行送り（フォントサイズ×行送り%）。ここに値を入れると % を逆算して設定します。",
                en: "The effective leading (font size × leading %). Enter a value here to set the leading by back-calculating the %."
            },
            reload: {
                ja: "選択中のテキストの現在値を読み取り直して UI に反映します。",
                en: "Re-read the current values from the selection and reflect them in the UI."
            },
            leadingAuto: {
                ja: "現在の行送り（絶対値）からフォントサイズに対する％を計算し、行送り（%）に反映して適用します。",
                en: "Compute the % of the current (absolute) leading relative to the font size, set it in Leading (%), and apply."
            },
            proportionalMetrics: {
                ja: "プロポーショナルメトリクス（和文プロポーショナル字形／メトリクス由来のツメ）をオン／オフします。",
                en: "Toggle proportional metrics (proportional CJK glyph spacing from font metrics)."
            }
        },
        alert: {
            applyError: { ja: "適用に失敗しました", en: "Apply failed" }
        }
    };

    /**
     * 表示言語のラベル文字列を返す
     * @param {Object} labelEntry - { ja, en } 形式のラベル定義
     * @returns {string} 表示言語のラベル（無ければ ja → en → 空文字）
     */
    function getLabel(labelEntry) {
        if (!labelEntry) return "";
        return labelEntry[uiLang] || labelEntry.ja || labelEntry.en || "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelEntry - ラベル定義
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelEntry) {
        return getLabel(labelEntry) + (uiLang === "ja" ? "：" : ":");
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

    /**
     * 行送りの単位ラベルを返す（文字サイズが Q のとき、行送りは「H」と呼ぶ）
     * @param {{code: number, label: string}} textUnit - 文字サイズの単位
     * @returns {string} 行送りの単位ラベル
     */
    function getLeadingUnitLabel(textUnit) {
        return (textUnit.code === 5) ? "H" : textUnit.label;
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // 以下の関数群は toString() で連結し、BridgeTalk 本文に同梱して
    // メインエンジン（生きた DOM を持つ）で eval される。常駐エンジン側
    // からは直接呼ばず、本文への埋め込み用途のみ。
    // toString() で送るため JSDoc は付けず、関数内のコメントも /* */ だけにする
    // These are serialized with toString(), so they carry no JSDoc and only /* */ comments
    // =========================================

    /* 型名を安全に取得 / Safely resolve a type name */
    function getTypeName(domObject) {
        if (domObject === null || domObject === undefined) return "";
        if (domObject.typename) return domObject.typename;
        try {
            return domObject.constructor ? domObject.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 例外を短いメッセージ文字列へ / Reduce an exception to a short message string */
    function errMessage(e) {
        return (e && e.message) ? e.message : String(e);
    }

    /* 1 アイテムからテキスト範囲を収集（グループは中を再帰）/ Collect text ranges from one item (descend into groups) */
    function collectTextRangesFromItem(pageItem, selectedRanges) {
        if (!pageItem) return;
        var itemType = getTypeName(pageItem);
        if (itemType === "TextFrame") {
            selectedRanges.push(pageItem.textRange);
        } else if (itemType === "TextRange") {
            selectedRanges.push(pageItem);
        } else if (itemType === "GroupItem" && pageItem.pageItems) {
            for (var i = 0; i < pageItem.pageItems.length; i++) collectTextRangesFromItem(pageItem.pageItems[i], selectedRanges);
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
    function getSelectedTextRanges() {
        var currentSelection = app.activeDocument.selection;
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
        var textRanges = getSelectedTextRanges();
        var paragraphRanges = [];
        for (var i = 0; i < textRanges.length; i++) {
            try {
                var paragraphs = textRanges[i].paragraphs;
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
    function applyKerningToRanges(textRanges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < textRanges.length; i++) {
            try {
                textRanges[i].characterAttributes.kerningMethod = kerningMethod;
                textRanges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                /* 適用できない範囲はスキップ / Skip ranges that can't take these attributes */
            }
        }
    }

    /* 選択範囲にプロポーショナルメトリクスを適用 / Apply proportional metrics to the given ranges */
    function applyPropMetricsToRanges(textRanges, useProportionalMetrics) {
        for (var i = 0; i < textRanges.length; i++) {
            try {
                textRanges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                /* 適用できない範囲はスキップ / Skip ranges that can't take this attribute */
            }
        }
    }

    /* 選択範囲に文字ツメを適用 / Apply Tsume to the given ranges（0〜100 の百分率）*/
    function applyTsumeToRanges(textRanges, tsumePercent) {
        for (var i = 0; i < textRanges.length; i++) {
            try {
                textRanges[i].characterAttributes.Tsume = tsumePercent;
            } catch (e) {
                /* 適用できない範囲はスキップ / Skip ranges that can't take this attribute */
            }
        }
    }

    /* 選択範囲にトラッキングを適用 / Apply tracking to the given ranges（1/1000 em 単位）*/
    function applyTrackingToRanges(textRanges, trackingValue) {
        for (var i = 0; i < textRanges.length; i++) {
            try {
                textRanges[i].characterAttributes.tracking = trackingValue;
            } catch (e) {
                /* 適用できない範囲はスキップ / Skip ranges that can't take this attribute */
            }
        }
    }

    /* 自動カーニング method を ID 文字列へ / Convert a kerning method (AutoKernType) to an id string */
    function kernMethodToId(kerningMethod) {
        var methodText = String(kerningMethod);
        if (methodText === String(AutoKernType.AUTO)) return "metrics";
        if (methodText === String(AutoKernType.OPTICAL)) return "optical";
        if (methodText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (methodText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /* テキスト範囲っぽい型か / Is this a text-range-like type */
    function isTextRangeLikeType(typeName) {
        return typeName === "TextRange" || typeName === "InsertionPoint" || typeName === "Character" ||
            typeName === "Word" || typeName === "Line" || typeName === "Paragraph";
    }

    /* 親をたどって TextFrame を返す / Walk up parents to the enclosing TextFrame */
    function findParentTextFrame(textItem) {
        var currentItem = textItem;
        for (var i = 0; i < 20; i++) {
            if (!currentItem) return null;
            if (getTypeName(currentItem) === "TextFrame") return currentItem;
            try { currentItem = currentItem.parent; } catch (e) { return null; }
        }
        return null;
    }

    /* TextFrame の重複判定キー / A dedup key for a TextFrame */
    function getTextFrameKey(textFrame) {
        try { if (textFrame.uuid) return textFrame.uuid; } catch (e) { }
        var bounds = textFrame.visibleBounds;
        var frameContents = "";
        try { frameContents = String(textFrame.contents); } catch (eContents) { }
        var contentKey = frameContents.length + ":" + frameContents.substring(0, 16);
        return textFrame.typename + ":" + textFrame.position[0] + ":" + textFrame.position[1] + ":" + bounds.join(":") + ":" + contentKey;
    }

    /* 処理可能なテキストフレームか判定し、該当すれば返す / Return the item if it is a processable TextFrame */
    function getProcessableTextFrame(pageItem) {
        if (!pageItem || pageItem.typename !== "TextFrame") return null;
        if (!pageItem.contents) return null;
        if (!pageItem.lines || pageItem.lines.length === 0) return null;
        return pageItem;
    }

    /* 選択から重複なしの処理対象テキストフレームを収集 / Collect unique processable text frames from the selection */
    function collectLeadingFrames(selectionItems) {
        var textFrames = [];
        var seenFrameKeys = {};
        var itemList = (selectionItems && selectionItems.typename) ? [selectionItems] : (selectionItems || []);
        for (var i = 0; i < itemList.length; i++) collectLeadingFramesFromItem(itemList[i], textFrames, seenFrameKeys);
        return textFrames;
    }

    /* 1 アイテムからフレームを収集（グループは中を走査、テキスト編集モードの範囲は親フレームへ）/ Collect frames from one item */
    function collectLeadingFramesFromItem(pageItem, textFrames, seenFrameKeys) {
        if (!pageItem) return;
        var typeName = getTypeName(pageItem);
        if (typeName === "TextFrame") { addLeadingFrame(pageItem, textFrames, seenFrameKeys); return; }
        if (isTextRangeLikeType(typeName)) { addLeadingFrame(findParentTextFrame(pageItem), textFrames, seenFrameKeys); return; }
        if (typeName === "GroupItem" && pageItem.pageItems) {
            for (var i = 0; i < pageItem.pageItems.length; i++) collectLeadingFramesFromItem(pageItem.pageItems[i], textFrames, seenFrameKeys);
        }
    }

    /* 処理可能なフレームを重複なく追加 / Add a processable frame, skipping duplicates */
    function addLeadingFrame(candidateFrame, textFrames, seenFrameKeys) {
        var textFrame = getProcessableTextFrame(candidateFrame);
        if (!textFrame) return;
        var frameKey = getTextFrameKey(textFrame);
        if (seenFrameKeys[frameKey]) return;
        seenFrameKeys[frameKey] = true;
        textFrames.push(textFrame);
    }

    /* 段落に自動行送りを設定：自動行送り量（%）を段落属性に、各文字を autoLeading=true に
       これにより Illustrator 上は常に「自動」表示となり、行送りはフォントサイズに追従する */
    function applyAutoLeadingToParagraphs(paragraphRanges, percent) {
        for (var i = 0; i < paragraphRanges.length; i++) {
            try {
                paragraphRanges[i].paragraphAttributes.autoLeadingAmount = percent;
                paragraphRanges[i].characterAttributes.autoLeading = true;
            } catch (e) {
                /* 適用できない範囲はスキップ / Skip ranges that can't take these attributes */
            }
        }
    }

    /* 選択の現在値を読み取り、初期表示用にエンコード / Read the selection's current state for reflection
       戻り値: count|fontSizePt|autoAmount|kernId|tsume|tracking|propMetrics|leadingPt
       注意: 文字属性は「最初のフレームの先頭文字」を代表値として読む簡略化。
       行送り%（autoAmount）は適用先と揃えるため、選択が触れている先頭段落から読む */
    function readState(textFrames, paragraphRanges) {
        var count = textFrames.length;
        var fontSizePt = NaN, autoAmount = NaN, kernId = "", tsume = NaN, tracking = NaN, propMetrics = 0, leadingPt = NaN;
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var lines = textFrames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var charAttributes = lines[0].characters[0].characterAttributes;
                    fontSizePt = charAttributes.size;
                    leadingPt = charAttributes.leading;
                    kernId = kernMethodToId(charAttributes.kerningMethod);
                    tsume = charAttributes.Tsume;
                    tracking = charAttributes.tracking;
                    propMetrics = charAttributes.proportionalMetrics ? 1 : 0;
                    break;
                }
            } catch (e) { }
        }
        if (paragraphRanges && paragraphRanges.length > 0) {
            try { autoAmount = paragraphRanges[0].paragraphAttributes.autoLeadingAmount; } catch (eAuto) { }
        }
        return [count, fontSizePt, autoAmount, kernId, tsume, tracking, propMetrics, leadingPt].join("|");
    }

    /* 選択の実際の行送り（絶対値 pt）とフォントサイズを読む / Read the actual leading (absolute pt) and font size
       戻り値: leadingPt|sizePt（先頭フレームの先頭文字を代表値として読む）*/
    function readLeadingAbs(textFrames) {
        var leadingPt = NaN, sizePt = NaN;
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var lines = textFrames[i].lines;
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

    /* 選択範囲へ作用する定型アクションの共通処理（空チェック→try→redraw→件数返却）/ Shared wrapper for range actions */
    function runRangeAction(textRanges, applyToRanges) {
        if (textRanges.length === 0) return "OK:0";
        try { applyToRanges(); } catch (e) { return "ERR:" + errMessage(e); }
        app.redraw();
        return "OK:" + textRanges.length;
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    /* メインエンジンへ送る処理関数（上で定義済みのものを再利用）
       Functions shipped to the main engine (reuse of the helpers above) */
    var WORKER_FUNCS = [
        getTypeName, errMessage, runRangeAction, collectTextRangesFromItem, getSelectedTextRanges, getSelectedParagraphRanges,
        resolveAutoKernType, applyKerningToRanges, applyPropMetricsToRanges, applyTsumeToRanges, applyTrackingToRanges, kernMethodToId,
        isTextRangeLikeType, findParentTextFrame, getTextFrameKey,
        getProcessableTextFrame, collectLeadingFrames, collectLeadingFramesFromItem, addLeadingFrame,
        applyAutoLeadingToParagraphs, readState, readLeadingAbs
    ];

    var workerLibCache = null;

    /**
     * 送る関数を toString() で連結したソース文字列を返す（初回だけ組み立てる）
     * @returns {string} メインエンジンへ同梱するソース
     */
    function getWorkerLibSource() {
        if (workerLibCache === null) {
            workerLibCache = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) workerLibCache += WORKER_FUNCS[i].toString() + "\n";
        }
        return workerLibCache;
    }

    /* メインエンジンで実行されるディスパッチャ / Dispatcher executed on the main engine
       結果は "OK:<payload>" または "ERR:<msg>" の文字列で返す（toString() で送るので JSDoc なし） */
    function dispatchAction(actionId, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        var textRanges = getSelectedTextRanges();
        if (actionId === "count") {
            return "OK:" + textRanges.length;
        }
        if (actionId === "getState") {
            return "OK:" + readState(collectLeadingFrames(app.activeDocument.selection), getSelectedParagraphRanges());
        }
        if (actionId === "getLeadingAbs") {
            return "OK:" + readLeadingAbs(collectLeadingFrames(app.activeDocument.selection));
        }
        /* 自動カーニング・文字ツメは段落単位（選択が触れた段落全体へ）/ Auto-kerning & Tsume apply per paragraph */
        if (actionId === "apply") { var kernParas = getSelectedParagraphRanges(); return runRangeAction(kernParas, function () { applyKerningToRanges(kernParas, resolveAutoKernType(params.method)); }); }
        if (actionId === "applyPropMetrics") { var propParas = getSelectedParagraphRanges(); return runRangeAction(propParas, function () { applyPropMetricsToRanges(propParas, params.propMetrics === 1); }); }
        if (actionId === "applyTsume") { var tsumeParas = getSelectedParagraphRanges(); return runRangeAction(tsumeParas, function () { applyTsumeToRanges(tsumeParas, params.value); }); }
        /* トラッキングはラン（選択）単位のまま / Tracking stays per selection (run) */
        if (actionId === "applyTracking") return runRangeAction(textRanges, function () { applyTrackingToRanges(textRanges, params.tracking); });
        if (actionId === "applyFontSize") return runRangeAction(textRanges, function () {
            for (var rangeIndex = 0; rangeIndex < textRanges.length; rangeIndex++) textRanges[rangeIndex].characterAttributes.size = params.sizePt;
        });
        if (actionId === "applyLeading") {
            var leadingFrames = collectLeadingFrames(app.activeDocument.selection);
            if (leadingFrames.length === 0) return "OK:0";
            try {
                /* 自動行送り量（%）を段落ごとに設定し、常に自動行送りに / Set the auto-leading amount (%) per paragraph; always auto-leading
                   行送りの基準（leadingType）はUIで扱わないため触らない / The leading basis is not exposed in the UI, so leave it alone */
                applyAutoLeadingToParagraphs(getSelectedParagraphRanges(), params.percent);
            } catch (errLeading) {
                return "ERR:" + errMessage(errLeading);
            }
            app.redraw();
            return "OK:" + leadingFrames.length;
        }
        return "OK:0";
    }
    var DISPATCH_SRC = "(" + dispatchAction.toString() + ")";

    /**
     * パラメータを安全な JS リテラル文字列にする
     * @param {Object|null} params - 送るパラメータ
     * @returns {string} オブジェクトリテラルのソース
     */
    function paramsToSource(params) {
        if (!params) return "{}";
        var literalParts = [];
        if (params.method !== undefined) literalParts.push('method:decodeURIComponent("' + encodeURIComponent(params.method) + '")');
        if (params.value !== undefined) literalParts.push("value:" + parseInt(params.value, 10));
        if (params.propMetrics !== undefined) literalParts.push("propMetrics:" + (params.propMetrics ? 1 : 0));
        if (params.tracking !== undefined) literalParts.push("tracking:" + parseInt(params.tracking, 10));
        if (params.sizePt !== undefined) literalParts.push("sizePt:" + params.sizePt);
        if (params.percent !== undefined) literalParts.push("percent:" + params.percent);
        return "{" + literalParts.join(",") + "}";
    }

    /**
     * ソースをメインエンジンへ送り、結果マーカーを解析して onDone(status, payload) を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープするため、コード全体を
     * encodeURIComponent で包んで送り、ターゲットで decodeURIComponent + eval して復元する
     * @param {string} workerCode - 実行するソース
     * @param {Function} onDone - 完了時に呼ぶ関数（"ok" / "error" と本文）
     * @returns {void}
     */
    function sendWorker(workerCode, onDone) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
        bridgeTalk.onResult = function (response) {
            var payload = response.body || "";
            var colonIndex = payload.indexOf(":");
            var marker = colonIndex >= 0 ? payload.substring(0, colonIndex) : payload;
            var rest = colonIndex >= 0 ? payload.substring(colonIndex + 1) : "";
            if (marker === "OK") onDone("ok", rest);
            else onDone("error", rest);
        };
        bridgeTalk.onError = function (response) { onDone("error", response && response.body ? response.body : "BridgeTalk error"); };
        bridgeTalk.send();
    }

    /**
     * メインエンジンへアクションを委譲する（非同期）
     * @param {string} actionId - dispatchAction() に渡すアクション名
     * @param {Object|null} params - パラメータ
     * @param {Function} onDone - 完了時に呼ぶ関数
     * @returns {void}
     */
    function runWorker(actionId, params, onDone) {
        var workerCode = getWorkerLibSource() + "\nvar __r=" + DISPATCH_SRC + "(\"" + actionId + "\"," + paramsToSource(params) + ");__r;";
        sendWorker(workerCode, onDone);
    }

    // =========================================
    // UI部品 / UI helpers
    // =========================================

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定（row / column で整列を切り替え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（既定）
     * @param {number} [spacing] - 間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 入力欄を ↑↓ キーで増減できるようにする（Shift で 10、Option で 0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} onUpdate - 値が変わったときに呼ぶ関数
     * @param {number} [decimals] - 表示する小数桁数
     * @returns {void}
     */
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

    /**
     * スライダー値を丸める（Shift 併用で 10 刻み）
     * @param {number} sliderValue - スライダーの値
     * @returns {number} 丸めた値
     */
    function snapSliderValue(sliderValue) {
        var isShiftPressed = ScriptUI.environment.keyboardState.shiftKey;
        return isShiftPressed ? Math.round(sliderValue / 10) * 10 : Math.round(sliderValue);
    }

    /**
     * スライダーを入力欄と連動させ、確定時に onCommit を呼ぶ
     * @param {Slider} slider - 対象のスライダー
     * @param {EditText} valueInput - 連動する入力欄
     * @param {Function} onCommit - 確定した値を受け取る関数
     * @returns {void}
     */
    function bindSliderToInput(slider, valueInput, onCommit) {
        function syncFromSlider() {
            var snappedValue = snapSliderValue(slider.value);
            slider.value = snappedValue;
            valueInput.text = String(snappedValue);
            return snappedValue;
        }
        slider.onChanging = syncFromSlider;
        slider.onChange = function () { onCommit(syncFromSlider()); };
    }

    /**
     * 入力欄を範囲内に丸めてスライダーへ反映し、onCommit を呼ぶハンドラを作る
     * @param {EditText} valueInput - 対象の入力欄
     * @param {Slider} slider - 連動するスライダー
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @param {Function} onCommit - 確定した値を受け取る関数
     * @returns {Function} 入力欄の onChange に設定するハンドラ
     */
    function makeClampedInputHandler(valueInput, slider, minValue, maxValue, onCommit) {
        return function () {
            var clampedValue = Math.round(parseFloat(valueInput.text));
            if (isNaN(clampedValue)) return;
            if (clampedValue < minValue) clampedValue = minValue;
            else if (clampedValue > maxValue) clampedValue = maxValue;
            valueInput.text = String(clampedValue);
            slider.value = clampedValue;
            onCommit(clampedValue);
        };
    }

    /**
     * スライダーと入力欄へ値を反映する
     * @param {Slider} slider - 対象のスライダー
     * @param {EditText} valueInput - 連動する入力欄
     * @param {number} value - 反映する値
     * @returns {void}
     */
    function reflectSliderValue(slider, valueInput, value) {
        slider.value = value;
        valueInput.text = String(value);
    }

    /**
     * ラジオを1つだけ選択する（別コンテナのラジオも手動で排他にする）
     * @param {RadioButton[]} radios - 対象のラジオ
     * @param {number} selectedIndex - 選択する位置
     * @returns {void}
     */
    function selectExclusiveRadio(radios, selectedIndex) {
        for (var i = 0; i < radios.length; i++) radios[i].value = (i === selectedIndex);
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /**
     * 自動カーニングの選択肢を返す
     * @returns {Array<{id: string, label: Object, tooltip: Object}>} 選択肢
     */
    function createAutoKernOptions() {
        return [
            { id: "mono", label: LABELS.radio.kernMono, tooltip: LABELS.tooltip.kernMono },
            { id: "zero", label: LABELS.radio.kernZero, tooltip: LABELS.tooltip.kernZero },
            { id: "metrics", label: LABELS.radio.kernMetrics, tooltip: LABELS.tooltip.kernMetrics },
            { id: "optical", label: LABELS.radio.kernOptical, tooltip: LABELS.tooltip.kernOptical }
        ];
    }

    /**
     * 「項目名＋入力欄＋単位」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelEntry - 項目名のラベル定義
     * @param {string} unitText - 単位の表示
     * @param {Object} tooltipEntry - 項目名と入力欄の tooltip
     * @returns {{fieldRow: Group, rowLabel: StaticText, fieldInput: EditText}} 作成した行
     */
    function addLeadingFieldRow(parentPanel, labelEntry, unitText, tooltipEntry) {
        var fieldRow = parentPanel.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        var rowLabel = fieldRow.add("statictext", undefined, labelText(labelEntry));
        var fieldInput = fieldRow.add("edittext", undefined, "");
        fieldInput.characters = LEADING_FIELD_CHARS;
        fieldRow.add("statictext", undefined, unitText);
        rowLabel.helpTip = getLabel(tooltipEntry);
        fieldInput.helpTip = rowLabel.helpTip;
        return { fieldRow: fieldRow, rowLabel: rowLabel, fieldInput: fieldInput };
    }

    /**
     * 「フォントサイズと行送り」パネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {{code: number, label: string}} textUnit - 文字サイズの単位
     * @returns {{fontSizeInput: EditText, leadingEffectiveInput: EditText, leadingPercentInput: EditText, leadingAutoButton: Button}} 作成したコントロール
     */
    function buildLeadingPanel(parentGroup, textUnit) {
        var leadingPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.sizeAndLeading));
        setupPanel(leadingPanel, INNER_PANEL_SPACING);
        leadingPanel.alignChildren = "left";
        leadingPanel.helpTip = getLabel(LABELS.tooltip.leading);

        /* フォントサイズ / Font size */
        var fontSizeRow = addLeadingFieldRow(leadingPanel, LABELS.fieldLabel.fontSize, textUnit.label, LABELS.tooltip.fontSize);

        /* 実質（フォントサイズ×行送り% の結果。ここに入力すると % を逆算）/ Effective leading (size × %) */
        var effectiveRow = addLeadingFieldRow(leadingPanel, LABELS.fieldLabel.leading, getLeadingUnitLabel(textUnit), LABELS.tooltip.leadingEffective);

        /* 行送り（自動行送り量 %）/ Leading (auto-leading amount %) */
        var percentRow = addLeadingFieldRow(leadingPanel, LABELS.fieldLabel.leadingPercent, "%", LABELS.tooltip.leading);
        var leadingAutoButton = percentRow.fieldRow.add("button", undefined, getLabel(LABELS.button.autoCalc));
        leadingAutoButton.preferredSize.width = AUTO_BUTTON_WIDTH;
        leadingAutoButton.helpTip = getLabel(LABELS.tooltip.leadingAuto);

        /* ラベル幅を揃える / Unify label widths */
        fontSizeRow.rowLabel.preferredSize.width = LEADING_LABEL_WIDTH;
        effectiveRow.rowLabel.preferredSize.width = LEADING_LABEL_WIDTH;
        percentRow.rowLabel.preferredSize.width = LEADING_LABEL_WIDTH;

        return {
            fontSizeInput: fontSizeRow.fieldInput,
            leadingEffectiveInput: effectiveRow.fieldInput,
            leadingPercentInput: percentRow.fieldInput,
            leadingAutoButton: leadingAutoButton
        };
    }

    /**
     * 「自動カーニング」パネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Array<{id: string, label: Object, tooltip: Object}>} autoKernOptions - 選択肢
     * @returns {{kernRadios: RadioButton[]}} 作成したラジオ
     */
    function buildAutoKernPanel(parentGroup, autoKernOptions) {
        var autoKernPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.autoKern));
        setupPanel(autoKernPanel, INNER_PANEL_SPACING);
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.helpTip = getLabel(LABELS.tooltip.autoKern);

        var kernRadios = [];
        for (var i = 0; i < autoKernOptions.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLabel(autoKernOptions[i].label));
            kernRadio.helpTip = getLabel(autoKernOptions[i].tooltip);
            kernRadio.value = false;
            kernRadio.index = i;
            kernRadios.push(kernRadio);
        }
        return { kernRadios: kernRadios };
    }

    /**
     * 「項目名＋入力欄（＋単位）」の行と、その下のスライダーを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelEntry - 項目名のラベル定義
     * @param {Object} tooltipEntry - 入力欄の tooltip
     * @param {Object} sliderTooltipEntry - スライダーの tooltip
     * @param {string|null} unitText - 単位の表示（無ければ null）
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {{valueInput: EditText, slider: Slider}} 作成した入力欄とスライダー
     */
    function addSliderField(parentPanel, labelEntry, tooltipEntry, sliderTooltipEntry, unitText, minValue, maxValue) {
        var fieldRow = parentPanel.add("group");
        setupGroup(fieldRow, "row");
        fieldRow.add("statictext", undefined, labelText(labelEntry));
        var valueInput = fieldRow.add("edittext", undefined, "0");
        valueInput.characters = SPACING_FIELD_CHARS;
        valueInput.helpTip = getLabel(tooltipEntry);
        if (unitText) fieldRow.add("statictext", undefined, unitText);
        var slider = parentPanel.add("slider", undefined, 0, minValue, maxValue);
        slider.helpTip = getLabel(sliderTooltipEntry);
        return { valueInput: valueInput, slider: slider };
    }

    /**
     * 「字間調整」（プロポーショナルメトリクス・文字ツメ・トラッキング）パネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {{tsumeInput: EditText, tsumeSlider: Slider, trackingInput: EditText, trackingSlider: Slider, propMetricsCheckbox: Checkbox}} 作成したコントロール
     */
    function buildSpacingPanel(parentGroup) {
        var spacingPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.spacingAdjust));
        setupPanel(spacingPanel, INNER_PANEL_SPACING);

        /* プロポーショナルメトリクス（自動カーニングの「メトリクス」に連動する現在ロジックを可視化）
           Proportional metrics (surfaces the current logic tied to the "Metrics" kerning option) */
        var propMetricsRow = spacingPanel.add("group");
        setupGroup(propMetricsRow, "row");
        propMetricsRow.margins = PROP_METRICS_MARGINS;
        var propMetricsCheckbox = propMetricsRow.add("checkbox", undefined, getLabel(LABELS.checkbox.propMetrics));
        propMetricsCheckbox.helpTip = getLabel(LABELS.tooltip.proportionalMetrics);

        var tsumeField = addSliderField(spacingPanel, LABELS.fieldLabel.tsume, LABELS.tooltip.tsume, LABELS.tooltip.tsumeSlider, "%", TSUME_MIN, TSUME_MAX);

        /* 文字ツメとトラッキングの間に少し余白 / A little gap between Tsume and Tracking */
        var spacingGap = spacingPanel.add("group");
        spacingGap.preferredSize.height = 1;

        var trackingField = addSliderField(spacingPanel, LABELS.fieldLabel.tracking, LABELS.tooltip.tracking, LABELS.tooltip.trackingSlider, null, TRACKING_MIN, TRACKING_MAX);

        return {
            tsumeInput: tsumeField.valueInput,
            tsumeSlider: tsumeField.slider,
            trackingInput: trackingField.valueInput,
            trackingSlider: trackingField.slider,
            propMetricsCheckbox: propMetricsCheckbox
        };
    }

    /**
     * パレットを組み立てて参照を返す（イベントは未接続）
     * @param {Array<{id: string, label: Object, tooltip: Object}>} autoKernOptions - 自動カーニングの選択肢
     * @returns {Object} パレットとコントロールの参照
     */
    function createPaletteUI(autoKernOptions) {
        var palette = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = "fill";

        var textUnit = getUnitInfo("text/units");

        /* 1カラム：フォントサイズと行送り／自動カーニング／字間調整を縦に並べる
           Single column: font size & leading / auto-kerning / letter spacing, stacked vertically */
        var mainColumn = palette.add("group");
        mainColumn.orientation = "column";
        mainColumn.alignChildren = ["fill", "top"];
        mainColumn.spacing = MAIN_COLUMN_SPACING;

        var leadingControls = buildLeadingPanel(mainColumn, textUnit);
        var autoKernControls = buildAutoKernPanel(mainColumn, autoKernOptions);
        var spacingControls = buildSpacingPanel(mainColumn);

        /* フッター（右＝再読み込み）/ Footer: reload */
        var footerGroup = palette.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        var footerSpacer = footerGroup.add("statictext", undefined, "");
        footerSpacer.alignment = ["fill", "center"];
        var reloadButton = footerGroup.add("button", undefined, getLabel(LABELS.button.reload));
        reloadButton.alignment = ["right", "center"];
        reloadButton.helpTip = getLabel(LABELS.tooltip.reload);

        return {
            palette: palette,
            textUnit: textUnit,
            fontSizeInput: leadingControls.fontSizeInput,
            leadingEffectiveInput: leadingControls.leadingEffectiveInput,
            leadingPercentInput: leadingControls.leadingPercentInput,
            leadingAutoButton: leadingControls.leadingAutoButton,
            kernRadios: autoKernControls.kernRadios,
            tsumeInput: spacingControls.tsumeInput,
            tsumeSlider: spacingControls.tsumeSlider,
            trackingInput: spacingControls.trackingInput,
            trackingSlider: spacingControls.trackingSlider,
            propMetricsCheckbox: spacingControls.propMetricsCheckbox,
            reloadButton: reloadButton
        };
    }

    // =========================================
    // 状態の反映 / Reflecting the selection state
    // =========================================

    /**
     * "count|fontSizePt|autoAmount|kernId|tsume|tracking|propMetrics|leadingPt" を分解する
     * @param {string} encodedState - readState() の戻り値
     * @returns {{count: number, fontSizePt: number, autoAmount: number, kernId: string, tsume: number, tracking: number, propMetrics: boolean, leadingPt: number}} 分解した値
     */
    function parseState(encodedState) {
        var stateFields = String(encodedState || "").split("|");
        function toNumber(fieldText) { var parsed = parseFloat(fieldText); return isNaN(parsed) ? NaN : parsed; }
        return {
            count: parseInt(stateFields[0], 10) || 0,
            fontSizePt: toNumber(stateFields[1]),
            autoAmount: toNumber(stateFields[2]),
            kernId: stateFields[3] || "",
            tsume: toNumber(stateFields[4]),
            tracking: toNumber(stateFields[5]),
            propMetrics: parseInt(stateFields[6], 10) === 1,
            leadingPt: toNumber(stateFields[7])
        };
    }

    /**
     * 方式 ID が一致するラジオを選択する
     * @param {RadioButton[]} kernRadios - 自動カーニングのラジオ
     * @param {Array<{id: string}>} autoKernOptions - 選択肢
     * @param {string} targetId - 選択する方式 ID
     * @returns {void}
     */
    function selectKernById(kernRadios, autoKernOptions, targetId) {
        for (var i = 0; i < autoKernOptions.length; i++) {
            if (autoKernOptions[i].id === targetId) { selectExclusiveRadio(kernRadios, i); return; }
        }
    }

    // =========================================
    // イベント接続 / Wire palette events
    // =========================================

    /**
     * 適用要求をメインエンジンへ順に送るキューを作る
     * BridgeTalk は非同期なので、実行中に来た要求は最新の1件だけ保持し、完了後に投げ直す
     * @returns {Function} runApply(actionId, params, onDone)
     */
    function createApplyQueue() {
        var workerBusy = false;
        var pendingApply = null;

        /* 重要処理の失敗をダイアログで通知 / Surface an important-op failure via an alert */
        function showWorkerError(actionId, payload) {
            var detail = payload ? (": " + String(payload)) : "";
            alert("⚠ " + getLabel(LABELS.alert.applyError) + " [" + actionId + "]" + detail);
        }

        /* 実際に委譲し、完了後に保留中の要求があれば続けて投げる / Delegate now, then flush the pending request if any */
        function sendApply(actionId, params, onDone) {
            workerBusy = true;
            runWorker(actionId, params, function (status, payload) {
                workerBusy = false;
                if (status === "error") showWorkerError(actionId, payload);
                if (onDone) onDone(status, payload);
                var nextApply = pendingApply;
                pendingApply = null;
                if (nextApply) sendApply(nextApply.actionId, nextApply.params, nextApply.onDone);
            });
        }

        return function runApply(actionId, params, onDone) {
            if (workerBusy) { pendingApply = { actionId: actionId, params: params, onDone: onDone }; return; }
            sendApply(actionId, params, onDone);
        };
    }

    /**
     * 自動カーニングとプロポーショナルメトリクスのイベントを接続する
     * @param {Object} paletteControls - createPaletteUI() の結果
     * @param {Array<{id: string}>} autoKernOptions - 選択肢
     * @param {Function} runApply - 適用キュー
     * @returns {void}
     */
    function bindKerningEvents(paletteControls, autoKernOptions, runApply) {
        for (var i = 0; i < paletteControls.kernRadios.length; i++) {
            paletteControls.kernRadios[i].onClick = function () {
                var methodId = autoKernOptions[this.index].id;
                runApply("apply", { method: methodId });
                /* 現在ロジック：メトリクスのときだけプロポーショナルメトリクスON / Current logic: only "metrics" turns it on */
                paletteControls.propMetricsCheckbox.value = (methodId === "metrics");
            };
        }

        /* プロポーショナルメトリクス：チェックで独立して適用（段落単位）/ Proportional metrics: apply independently on toggle */
        paletteControls.propMetricsCheckbox.onClick = function () {
            runApply("applyPropMetrics", { propMetrics: this.value });
        };
    }

    /**
     * スライダー＋入力欄の組を、範囲内に丸めて適用するよう接続する（Shift 併用で 10 刻み）
     * @param {Slider} slider - スライダー
     * @param {EditText} valueInput - 入力欄
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @param {boolean} allowNegative - ↑↓キーで負の値を許すか
     * @param {Function} applyValue - 確定した値を適用する関数
     * @returns {void}
     */
    function bindSliderField(slider, valueInput, minValue, maxValue, allowNegative, applyValue) {
        bindSliderToInput(slider, valueInput, applyValue);
        var applyFromInput = makeClampedInputHandler(valueInput, slider, minValue, maxValue, applyValue);
        valueInput.onChange = applyFromInput;
        changeValueByArrowKey(valueInput, allowNegative, applyFromInput, 0);
    }

    /**
     * 文字ツメとトラッキングのイベントを接続する
     * @param {Object} paletteControls - createPaletteUI() の結果
     * @param {Function} runApply - 適用キュー
     * @returns {void}
     */
    function bindSpacingEvents(paletteControls, runApply) {
        bindSliderField(paletteControls.tsumeSlider, paletteControls.tsumeInput, TSUME_MIN, TSUME_MAX, false, function (tsumeValue) {
            runApply("applyTsume", { value: tsumeValue });
        });
        bindSliderField(paletteControls.trackingSlider, paletteControls.trackingInput, TRACKING_MIN, TRACKING_MAX, true, function (trackingValue) {
            runApply("applyTracking", { tracking: trackingValue });
        });
    }

    /**
     * フォントサイズと行送りのイベントを接続する
     * @param {Object} paletteControls - createPaletteUI() の結果
     * @param {Function} runApply - 適用キュー
     * @returns {void}
     */
    function bindLeadingEvents(paletteControls, runApply) {
        var textUnit = paletteControls.textUnit;

        /* 現在の行送り（自動行送り量 %）/ The current leading (auto-leading amount %) */
        function currentLeadingPercent() {
            return parseFloat(paletteControls.leadingPercentInput.text);
        }
        /* 実質行送り（フォントサイズ×%）の表示を更新 / Update the effective-leading display (font size × %) */
        function updateLeadingEffective() {
            var fontSize = parseFloat(paletteControls.fontSizeInput.text);
            var percent = currentLeadingPercent();
            if (isNaN(fontSize) || isNaN(percent)) { paletteControls.leadingEffectiveInput.text = ""; return; }
            paletteControls.leadingEffectiveInput.text = String(Math.round(fontSize * percent / 100 * 10) / 10);
        }
        /* 行送りを適用（自動行送り量%。行送りの基準は変更しない）/ Apply leading (auto-leading amount %; the basis is left as is) */
        function applyLeading() {
            var percent = currentLeadingPercent();
            if (isNaN(percent)) return;
            runApply("applyLeading", { percent: percent });
        }

        /* フォントサイズ入力：適用し、実質行送りの表示も更新。行送りは自動行送りなのでサイズに追従する（再適用は不要）
           Font size input: apply and refresh the effective leading; auto leading follows the size */
        function applyFontSizeFromInput() {
            var inputValue = parseFloat(paletteControls.fontSizeInput.text);
            if (!isNaN(inputValue)) runApply("applyFontSize", { sizePt: inputValue * textUnit.pointsPerUnit });
            updateLeadingEffective();
        }
        paletteControls.fontSizeInput.onChange = applyFontSizeFromInput;
        paletteControls.fontSizeInput.onChanging = function () { updateLeadingEffective(); };
        changeValueByArrowKey(paletteControls.fontSizeInput, false, applyFontSizeFromInput, 1);

        /* 行送り（%）入力：実質表示を更新して適用 / Leading (%) input: refresh the effective display and apply */
        function applyLeadingFromPercent() {
            updateLeadingEffective();
            applyLeading();
        }
        paletteControls.leadingPercentInput.onChange = applyLeadingFromPercent;
        paletteControls.leadingPercentInput.onChanging = function () { updateLeadingEffective(); };
        changeValueByArrowKey(paletteControls.leadingPercentInput, false, applyLeadingFromPercent, 1);

        /* 実質（pt）入力：フォントサイズから % を逆算して適用 / Effective (pt) input: back-calculate the % from the font size and apply */
        function applyLeadingFromEffective() {
            var effectiveLeading = parseFloat(paletteControls.leadingEffectiveInput.text);
            var fontSize = parseFloat(paletteControls.fontSizeInput.text);
            if (isNaN(effectiveLeading) || isNaN(fontSize) || fontSize <= 0) return;
            var percent = Math.round((effectiveLeading / fontSize) * 100 * 10) / 10;
            paletteControls.leadingPercentInput.text = String(percent);
            applyLeading();
        }
        paletteControls.leadingEffectiveInput.onChange = applyLeadingFromEffective;
        changeValueByArrowKey(paletteControls.leadingEffectiveInput, false, applyLeadingFromEffective, 1);

        /* 自動計算ボタン：現在の行送り（絶対値）から % を逆算して行送り（%）へ反映し適用
           Auto-calc button: back-calculate the % from the current (absolute) leading, set Leading (%), and apply */
        paletteControls.leadingAutoButton.onClick = function () {
            runWorker("getLeadingAbs", null, function (status, payload) {
                if (status !== "ok") return;
                var leadingParts = String(payload).split("|");
                var leadingPt = parseFloat(leadingParts[0]);
                var sizePt = parseFloat(leadingParts[1]);
                if (isNaN(leadingPt) || isNaN(sizePt) || sizePt <= 0) return;
                var percent = Math.round((leadingPt / sizePt) * 100 * 10) / 10;
                paletteControls.leadingPercentInput.text = String(percent);
                updateLeadingEffective();
                applyLeading();
            });
        };
    }

    /**
     * 選択状態を読み取って UI に反映する関数を作る
     * @param {Object} paletteControls - createPaletteUI() の結果
     * @param {Array<{id: string}>} autoKernOptions - 選択肢
     * @returns {Function} refreshState()
     */
    function makeStateRefresher(paletteControls, autoKernOptions) {
        var textUnit = paletteControls.textUnit;
        return function refreshState() {
            runWorker("getState", null, function (status, payload) {
                if (status !== "ok") return;
                var selectionState = parseState(payload);
                if (selectionState.count <= 0) return;
                /* フォントサイズ / Font size */
                paletteControls.fontSizeInput.text = isNaN(selectionState.fontSizePt) ? "" : String(Math.round((selectionState.fontSizePt / textUnit.pointsPerUnit) * 10) / 10);
                /* 行送り%：自動行送り量% を % 欄に反映（実質欄は下で現在値に上書きするので計算表示は使わない）
                   Leading %: reflect the auto-leading amount % (the effective field is overwritten by the actual value below) */
                paletteControls.leadingPercentInput.text = isNaN(selectionState.autoAmount) ? "" : String(Math.round(selectionState.autoAmount * 10) / 10);
                /* 行送り：選択の現在値（絶対値 pt）をそのまま表示（サイズ×% の計算値ではない）
                   Leading: show the selection's actual current value (absolute pt), not the size × % computation */
                paletteControls.leadingEffectiveInput.text = isNaN(selectionState.leadingPt) ? "" : String(Math.round((selectionState.leadingPt / textUnit.pointsPerUnit) * 10) / 10);
                /* 自動カーニング / Auto kerning */
                if (selectionState.kernId) selectKernById(paletteControls.kernRadios, autoKernOptions, selectionState.kernId);
                /* 文字ツメ / Tsume */
                if (!isNaN(selectionState.tsume)) reflectSliderValue(paletteControls.tsumeSlider, paletteControls.tsumeInput, Math.round(selectionState.tsume));
                /* トラッキング / Tracking */
                if (!isNaN(selectionState.tracking)) reflectSliderValue(paletteControls.trackingSlider, paletteControls.trackingInput, Math.round(selectionState.tracking));
                /* プロポーショナルメトリクス / Proportional metrics */
                paletteControls.propMetricsCheckbox.value = !!selectionState.propMetrics;
            });
        };
    }

    /**
     * パレットのイベントをすべて接続する
     * @param {Object} paletteControls - createPaletteUI() の結果
     * @param {Array<{id: string}>} autoKernOptions - 選択肢
     * @returns {{refreshState: Function}} 選択状態を反映する関数
     */
    function bindPaletteEvents(paletteControls, autoKernOptions) {
        var runApply = createApplyQueue();
        bindKerningEvents(paletteControls, autoKernOptions, runApply);
        bindSpacingEvents(paletteControls, runApply);
        bindLeadingEvents(paletteControls, runApply);

        var refreshState = makeStateRefresher(paletteControls, autoKernOptions);

        /* パレット表示・フォーカス復帰のたびに選択状況を反映 / Reflect the selection on show and on regaining focus */
        paletteControls.palette.onShow = function () { refreshState(); };
        paletteControls.palette.onActivate = function () { refreshState(); };

        /* 再読み込み：選択中のテキストの現在値を読み直して UI に反映 / Reload: re-read the selection's current values into the UI */
        paletteControls.reloadButton.onClick = function () { refreshState(); };

        /* Esc でパレットを閉じる / Close the palette on Esc */
        paletteControls.palette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") paletteControls.palette.close();
        });

        return { refreshState: refreshState };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 常駐エンジンに保持するパレット参照のキー / Global key holding the persistent palette instance */
    var PALETTE_GLOBAL_KEY = "__typeBasicsPanelInstance";

    /**
     * パレットを開く（開いていれば前面に出す）
     * @returns {void}
     */
    function main() {
        /* 二重起動防止：既に開いているパレットがあれば新規作成せず前面に出して終了
           Prevent double launch: if a palette is already open, bring it to the front and exit */
        var existingPalette = $.global[PALETTE_GLOBAL_KEY];
        if (existingPalette) {
            try {
                existingPalette.show();
                existingPalette.active = true;
                return;
            } catch (e) {
                /* 参照が無効（既に破棄済み）なら作り直す / Stale reference: fall through and rebuild */
                $.global[PALETTE_GLOBAL_KEY] = null;
            }
        }

        var autoKernOptions = createAutoKernOptions();
        var paletteControls = createPaletteUI(autoKernOptions);
        bindPaletteEvents(paletteControls, autoKernOptions);

        /* パレット参照を常駐エンジンに保持し、閉じたら解放 / Keep the instance on the resident engine; clear it on close */
        $.global[PALETTE_GLOBAL_KEY] = paletteControls.palette;
        paletteControls.palette.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        paletteControls.palette.show();
    }

    main();

})();
