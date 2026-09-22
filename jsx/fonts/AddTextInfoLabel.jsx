#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストオブジェクトの下または右に、フォント情報を表示するラベルを作成します。
詳細表示（エリア内文字）と簡易表示（ポイント文字）を切り替えられ、表示項目は個別に指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTextInfoLabel.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n607ef418877f

### Overview

Adds a label showing the font information below or beside the selected text objects.
You can switch between a detailed layout (area text) and a compact one (point text), and choose which items appear.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTextInfoLabel.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTextInfoLabel";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTextInfoLabel.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTextInfoLabel.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n607ef418877f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 情報ラベルを入れるレイヤー名 / Layer that receives the info labels */
    var INFO_LAYER_NAME = "フォント情報";

    /* 情報ラベルの書式 / Info label text format */
    var LABEL_FONT_NAME = "HiraginoSans-W3";
    var LABEL_FONT_SIZE = 10;
    var LABEL_LEADING = 16;

    /* 一度に処理する選択数の上限（これ以上は何もしない） / Selections at or above this count are ignored */
    var MAX_SELECTION_COUNT = 1000;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白 / Panel margins */
    var PANEL_MARGINS = [15, 20, 15, 15];

    /* 一括切り替えボタンの大きさ / Bounds of the bulk toggle buttons */
    var TOGGLE_BUTTON_BOUNDS = [0, 0, 80, 24];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキスト情報を追加", en: "Add Text Info Label" }
        },
        panel: {
            position: { ja: "位置", en: "Position" },
            mode: { ja: "表示形式", en: "Format" },
            info: { ja: "表示項目（詳細表示時のみ有効）", en: "Items (used by the detailed format only)" }
        },
        radio: {
            posBottom: { ja: "下", en: "Below" },
            posRight: { ja: "右", en: "Right" },
            modeCompact: { ja: "簡易版", en: "Compact" },
            modeFull: { ja: "詳細", en: "Detailed" }
        },
        checkbox: {
            fontName: { ja: "フォント名", en: "Font name" },
            postScript: { ja: "PSフォント名", en: "PostScript name" },
            fontStyle: { ja: "スタイル（ウェイト）", en: "Style (weight)" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            leading: { ja: "行送り", en: "Leading" },
            kerning: { ja: "カーニング", en: "Kerning" },
            proportional: { ja: "プロポーショナルメトリクス", en: "Proportional metrics" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            leadingPercent: { ja: "行送り（%）", en: "Leading (%)" }
        },
        tooltip: {
            posBottom: { ja: "テキストの下に情報ラベルを置きます。", en: "Places the info label below the text." },
            posRight: { ja: "テキストの右に情報ラベルを置きます。", en: "Places the info label to the right of the text." },
            modeCompact: { ja: "フォント名とサイズだけの短い表記にします。", en: "Writes a short label with just the font name and size." },
            modeFull: { ja: "下の［表示項目］で選んだ内容をすべて書き出します。", en: "Writes every item ticked under Items below." },
            info: { ja: "詳細表示のときに、ラベルへ書き出す項目を選びます。", en: "Picks which items go into the label when the detailed format is used." },
            kerning: {
                ja: "カーニングの方式（メトリクス／和文等幅／オプティカル／なし）を書き出します。",
                en: "Writes the kerning method (Metrics, Metrics - Roman Only, Optical, or none)."
            },
            leadingPercent: { ja: "行送りをフォントサイズに対する割合で書き出します。", en: "Writes the leading as a percentage of the font size." },
            minimalSet: {
                ja: "フォント名・スタイル・フォントサイズ・行送りだけをオンにします。",
                en: "Turns on only the font name, style, font size, and leading."
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.position" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 表示項目 / Info items
    // =========================================

    /* 詳細表示の項目（チェックボックスの列・初期値・［最小セット］での値）
       Detailed-format items: checkbox column, initial value, value set by the minimal preset */
    var INFO_ITEMS = [
        { key: "fontName",       column: 0, initial: true,  minimal: true },
        { key: "postScript",     column: 0, initial: false, minimal: false },
        { key: "fontStyle",      column: 0, initial: true,  minimal: true },
        { key: "fontSize",       column: 0, initial: true,  minimal: true },
        { key: "leading",        column: 0, initial: true,  minimal: true },
        { key: "kerning",        column: 1, initial: true,  minimal: false, tooltip: "tooltip.kerning" },
        { key: "proportional",   column: 1, initial: true,  minimal: false },
        { key: "tracking",       column: 1, initial: false, minimal: false },
        { key: "tsume",          column: 1, initial: false, minimal: false },
        { key: "leadingPercent", column: 1, initial: false, minimal: false, tooltip: "tooltip.leadingPercent" }
    ];

    // =========================================
    // テキスト情報の読み取り / Reading text attributes
    // =========================================

    /* 段落の区切り（CR） / Paragraph separator (CR) */
    var PARAGRAPH_BREAK = String.fromCharCode(13);

    /**
     * 環境設定の文字単位の表示名を返す
     * @param {string} prefKey - "text/units"（文字サイズ）または "text/asianunits"（行送り）
     * @returns {string} 単位の表示名（コード5は文字サイズなら Q、行送りなら H）
     */
    function getTextUnitLabel(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey);
        switch (unitCode) {
            case 0: return "inch";
            case 1: return "mm";
            case 2: return "pt";
            case 3: return "p";
            case 4: return "cm";
            case 5: return (prefKey === "text/units") ? "Q" : "H";
            case 6: return "px";
            default: return "pt";
        }
    }

    /**
     * 読み取り関数を呼び、例外が出たら「不明」を返す
     * @param {Function} readValue - (textFrame, displayMode) を受け取る読み取り関数
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} [displayMode] - "compact" または "full"
     * @returns {*} 読み取った値、または "不明"
     */
    function readOrUnknown(readValue, textFrame, displayMode) {
        /* 文字属性の読み取りは失敗しうる / reading character attributes may throw */
        try {
            return readValue(textFrame, displayMode);
        } catch (e) {
            return "不明";
        }
    }

    /**
     * フォントサイズを小数1桁の単位付き文字列で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} フォントサイズ
     */
    function getFontSize(textFrame) {
        var rawSize = textFrame.textRange.characterAttributes.size;
        var roundedSize = Math.round(rawSize * 10) / 10;
        return roundedSize + " " + getTextUnitLabel("text/units");
    }

    /**
     * 小数が3桁以上なら2桁に丸め、それ以外はそのまま文字列にする
     * @param {number} value - 数値
     * @returns {string} 整形した数値
     */
    function formatUpToTwoDecimals(value) {
        var valueText = String(value);
        if (valueText.indexOf(".") === -1) {
            return valueText;
        }
        if (valueText.split(".")[1].length <= 2) {
            return valueText;
        }
        return String(Math.round(value * 100) / 100);
    }

    /**
     * 行送りを単位付き文字列で返す（詳細表示の自動行送りは「自動（…）」）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} displayMode - "compact" または "full"
     * @returns {string} 行送り
     */
    function getLeading(textFrame, displayMode) {
        var textAttributes = textFrame.textRange.characterAttributes;
        var leadingText = formatUpToTwoDecimals(textAttributes.leading) + " " + getTextUnitLabel("text/asianunits");

        if (textAttributes.autoLeading && displayMode !== "compact") {
            return "自動（" + leadingText + "）";
        }
        return leadingText;
    }

    /**
     * 行送りのフォントサイズに対する割合を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} 「120 %」のような割合（数値が取れなければ「不明」）
     */
    function getLeadingPercentage(textFrame) {
        var textAttributes = textFrame.textRange.characterAttributes;
        var fontSize = textAttributes.size;
        var leading = textAttributes.leading;

        /* 数値でない場合は不明 / Unknown unless both are numbers */
        if (typeof fontSize !== "number" || typeof leading !== "number" || fontSize === 0) {
            return "不明";
        }
        return roundToTwoDecimalFixed((leading / fontSize) * 100) + " %";
    }

    /**
     * 小数2桁に丸め、「.00」なら整数にする
     * @param {number} value - 数値
     * @returns {string} 整形した数値
     */
    function roundToTwoDecimalFixed(value) {
        var rounded = Math.round(value * 100) / 100;
        var fixedText = rounded.toFixed(2);
        if (fixedText.match(/\.00$/)) return String(parseInt(rounded, 10));
        return fixedText;
    }

    /**
     * カーニングの方式を日本語名で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} カーニングの方式
     */
    function getKerningMethodText(textFrame) {
        switch (textFrame.textRange.characterAttributes.kerningMethod) {
            case AutoKernType.AUTO: return "メトリクス";
            case AutoKernType.METRICSROMANONLY: return "和文等幅";
            case AutoKernType.OPTICAL: return "オプティカル";
            default: return "なし";
        }
    }

    /**
     * トラッキングを返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number} トラッキング
     */
    function getTracking(textFrame) {
        return textFrame.textRange.characterAttributes.tracking;
    }

    /**
     * プロポーショナルメトリクスの状態を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} "ON" または "OFF"
     */
    function getProportionalMetrics(textFrame) {
        return textFrame.textRange.characterAttributes.proportionalMetrics ? "ON" : "OFF";
    }

    /**
     * 文字ツメを小数1桁までの文字列で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} 文字ツメ（数値が取れなければ「なし」）
     */
    function getTsume(textFrame) {
        var tsume = textFrame.textRange.characterAttributes.Tsume;
        if (typeof tsume !== "number" || isNaN(tsume)) return "なし";

        var rounded = Math.round(tsume * 10) / 10;
        return (rounded % 1 === 0) ? String(rounded.toFixed(0)) : String(rounded.toFixed(1));
    }

    /**
     * 最初に見つかったフォントを返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {TextFont|null} フォント（見つからなければ null）
     */
    function getFirstAvailableFont(textFrame) {
        var characters = textFrame.textRange.characters;
        for (var i = 0; i < characters.length; i++) {
            var textFont = characters[i].characterAttributes.textFont;
            if (textFont) return textFont;
        }
        return null;
    }

    /**
     * ラベルに書き出すテキスト情報をまとめて読み取る
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {TextFont} sourceFont - 対象のフォント
     * @param {string} displayMode - "compact" または "full"
     * @returns {Object} テキスト情報
     */
    function readTextInfo(textFrame, sourceFont, displayMode) {
        return {
            fontSize: readOrUnknown(getFontSize, textFrame),
            leading: readOrUnknown(getLeading, textFrame, displayMode),
            leadingPercent: readOrUnknown(getLeadingPercentage, textFrame),
            kerning: readOrUnknown(getKerningMethodText, textFrame),
            tracking: readOrUnknown(getTracking, textFrame),
            proportional: readOrUnknown(getProportionalMetrics, textFrame),
            tsume: readOrUnknown(getTsume, textFrame),
            fontFamily: sourceFont.family,
            fontStyle: sourceFont.style,
            postScriptName: sourceFont.name
        };
    }

    // =========================================
    // 情報ラベルの作成 / Info label creation
    // =========================================

    /**
     * 簡易表示の3行のテキストを組み立てる
     * @param {Object} textInfo - readTextInfo() の結果
     * @returns {string} ラベルのテキスト
     */
    function buildCompactText(textInfo) {
        var fontLine = textInfo.fontFamily + " " + textInfo.fontStyle + "、" + textInfo.fontSize + " ↓" + textInfo.leading;
        var kerningLine = textInfo.kerning + "、プロポーショナルメトリクス：" + textInfo.proportional;
        var spacingLine = "トラッキング：" + textInfo.tracking + "、文字ツメ：" + textInfo.tsume + " %";
        return fontLine + PARAGRAPH_BREAK + kerningLine + PARAGRAPH_BREAK + spacingLine;
    }

    /**
     * 詳細表示のテキストを、選ばれた項目だけで組み立てる
     * @param {Object} textInfo - readTextInfo() の結果
     * @param {Object} includeItems - 項目キーごとの書き出す／書き出さない
     * @returns {string} ラベルのテキスト
     */
    function buildDetailedText(textInfo, includeItems) {
        /* 書き出す順 / Output order */
        var detailLines = [
            ["fontName", "・フォント名\t" + textInfo.fontFamily],
            ["postScript", "・PSフォント名\t" + textInfo.postScriptName],
            ["fontStyle", "・スタイル（ウェイト）\t" + textInfo.fontStyle],
            ["fontSize", "・フォントサイズ\t" + textInfo.fontSize],
            ["leading", "・行送り\t" + textInfo.leading],
            ["leadingPercent", "・行送り（%）\t" + textInfo.leadingPercent],
            ["kerning", "・カーニング\t" + textInfo.kerning],
            ["proportional", "・プロポーショナルメトリクス\t" + textInfo.proportional],
            ["tracking", "・トラッキング\t" + textInfo.tracking],
            ["tsume", "・文字ツメ\t" + textInfo.tsume + " %"]
        ];
        var infoLines = [];
        for (var i = 0; i < detailLines.length; i++) {
            if (includeItems[detailLines[i][0]]) infoLines.push(detailLines[i][1]);
        }
        return infoLines.join(PARAGRAPH_BREAK);
    }

    /**
     * 詳細表示のラベルをエリア内文字で作る（右揃えのタブとリーダー付き）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - ラベルのテキスト
     * @param {number[]} sourceBounds - 元のテキストの geometricBounds
     * @param {string} position - "bottom" または "right"
     * @returns {TextFrame} 作ったラベル
     */
    function createDetailedInfoFrame(doc, textContent, sourceBounds, position) {
        var mmToPt = 72 / 25.4;
        var lineCount = textContent.split(PARAGRAPH_BREAK).length;
        var lineHeight = LABEL_FONT_SIZE * 1.4;
        var frameHeight = lineHeight * (lineCount + 2) + 4 * mmToPt;
        var frameWidth = 300;

        var frameLeft = (position === "right") ? sourceBounds[2] + 10 : sourceBounds[0];
        var frameTop = (position === "right") ? sourceBounds[3] + frameHeight : sourceBounds[3] - LABEL_FONT_SIZE;

        var framePath = doc.pathItems.rectangle(frameTop, frameLeft, frameWidth, frameHeight);
        var infoFrame = doc.textFrames.areaText(framePath);
        infoFrame.contents = textContent;
        infoFrame.spacing = 2 * mmToPt;

        /* タブストップ（右揃え、位置400pt、リーダー…） / Right tab stop at 400pt with a "…" leader */
        var tabStop = new TabStopInfo();
        tabStop.position = 400;
        tabStop.alignment = TabStopAlignment.Right;
        tabStop.leader = "…";
        for (var i = 0; i < infoFrame.paragraphs.length; i++) {
            infoFrame.paragraphs[i].tabStops = [tabStop];
        }
        return infoFrame;
    }

    /**
     * 簡易表示のラベルをポイント文字で作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - ラベルのテキスト
     * @param {number[]} sourceBounds - 元のテキストの geometricBounds
     * @param {string} position - "bottom" または "right"
     * @returns {TextFrame} 作ったラベル
     */
    function createCompactInfoFrame(doc, textContent, sourceBounds, position) {
        var labelLeft = (position === "right") ? sourceBounds[2] + 20 : sourceBounds[0];
        var labelTop = (position === "right") ? sourceBounds[1] : sourceBounds[3] - LABEL_FONT_SIZE;

        var infoFrame = doc.textFrames.add();
        infoFrame.contents = textContent;
        infoFrame.position = [labelLeft, labelTop];
        /* 見た目の上端をそろえる / Align the visible top edge */
        var infoBounds = infoFrame.visibleBounds;
        infoFrame.translate(0, labelTop - infoBounds[1]);
        return infoFrame;
    }

    /**
     * ラベルに文字サイズ・行送り・揃え・フォントを設定する
     * @param {TextFrame} infoFrame - 作ったラベル
     * @param {string} displayMode - "compact" または "full"
     * @returns {void}
     */
    function applyInfoFrameStyle(infoFrame, displayMode) {
        var infoAttributes = infoFrame.textRange.characterAttributes;
        infoAttributes.autoLeading = false;
        infoAttributes.leading = LABEL_LEADING;
        infoFrame.textRange.characterAttributes.size = LABEL_FONT_SIZE;
        infoFrame.textRange.paragraphAttributes.justification =
            (displayMode === "full") ? Justification.RIGHT : Justification.LEFT;

        /* フォントが無ければ既定のまま / keep the default font when it is missing */
        try {
            infoFrame.textRange.characterAttributes.textFont = textFonts.getByName(LABEL_FONT_NAME);
        } catch (e) {}
    }

    /**
     * 名前でレイヤーを探し、無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} レイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタン2つを横に並べたパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} panelLabelPath - パネルタイトルの LABELS パス
     * @param {string[]} radioKeys - LABELS.radio / LABELS.tooltip のキー
     * @returns {Object} キーごとのラジオボタン
     */
    function addRadioPanel(parentGroup, panelLabelPath, radioKeys) {
        var radioPanel = parentGroup.add("panel", undefined, getLabel(panelLabelPath));
        radioPanel.orientation = "row";
        radioPanel.margins = PANEL_MARGINS;
        var radios = {};
        for (var i = 0; i < radioKeys.length; i++) {
            var radio = radioPanel.add("radiobutton", undefined, getLabel("radio." + radioKeys[i]));
            radio.helpTip = getLabel("tooltip." + radioKeys[i]);
            radios[radioKeys[i]] = radio;
        }
        return radios;
    }

    /**
     * 表示項目パネル（2列のチェックボックスと一括切り替えボタン）を追加する
     * @param {Window} optionDialog - ダイアログ
     * @returns {{panel: Panel, checkboxes: Object}} パネルと項目キーごとのチェックボックス
     */
    function addInfoItemsPanel(optionDialog) {
        var infoPanel = optionDialog.add("panel", undefined, getLabel("panel.info"));
        infoPanel.helpTip = getLabel("tooltip.info");
        infoPanel.orientation = "column";
        infoPanel.alignChildren = "left";
        infoPanel.margins = PANEL_MARGINS;

        var columnsGroup = infoPanel.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        var columns = [];
        for (var c = 0; c < 2; c++) {
            columns[c] = columnsGroup.add("group");
            columns[c].orientation = "column";
            columns[c].alignChildren = "left";
        }

        var checkboxes = {};
        for (var i = 0; i < INFO_ITEMS.length; i++) {
            var infoItem = INFO_ITEMS[i];
            var itemCheckbox = columns[infoItem.column].add("checkbox", undefined, getLabel("checkbox." + infoItem.key));
            if (infoItem.tooltip) itemCheckbox.helpTip = getLabel(infoItem.tooltip);
            itemCheckbox.value = infoItem.initial;
            checkboxes[infoItem.key] = itemCheckbox;
        }

        /**
         * 全項目のチェックを、項目ごとの値で設定する
         * @param {Function} getValue - 項目定義を受け取って値を返す関数
         * @returns {void}
         */
        function setAllItems(getValue) {
            for (var i = 0; i < INFO_ITEMS.length; i++) {
                checkboxes[INFO_ITEMS[i].key].value = getValue(INFO_ITEMS[i]);
            }
        }

        var toggleButtonGroup = infoPanel.add("group");
        toggleButtonGroup.orientation = "row";
        toggleButtonGroup.alignment = "left";
        var btnAllOn = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "すべてON");
        var btnAllOff = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "すべてOFF");
        var btnMinimal = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "最小セット");
        btnMinimal.helpTip = getLabel("tooltip.minimalSet");

        btnAllOn.onClick = function () {
            setAllItems(function () { return true; });
        };
        btnAllOff.onClick = function () {
            setAllItems(function () { return false; });
        };
        btnMinimal.onClick = function () {
            setAllItems(function (infoItem) { return infoItem.minimal; });
        };

        return { panel: infoPanel, checkboxes: checkboxes };
    }

    /**
     * 設定ダイアログを表示し、選ばれた設定を返す
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showOptionDialog() {
        var optionDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionDialog.alignChildren = "left";

        var topGroup = optionDialog.add("group");
        topGroup.orientation = "row";
        topGroup.alignChildren = "top";

        var positionRadios = addRadioPanel(topGroup, "panel.position", ["posBottom", "posRight"]);
        positionRadios.posRight.value = true;

        var modeRadios = addRadioPanel(topGroup, "panel.mode", ["modeCompact", "modeFull"]);
        modeRadios.modeCompact.value = true;

        var infoItemsUI = addInfoItemsPanel(optionDialog);

        /* 表示項目は詳細表示のときだけ有効 / Items apply to the detailed format only */
        function updateInfoPanelEnabled() {
            infoItemsUI.panel.enabled = modeRadios.modeFull.value;
        }
        modeRadios.modeFull.onClick = updateInfoPanelEnabled;
        modeRadios.modeCompact.onClick = updateInfoPanelEnabled;
        updateInfoPanelEnabled();

        var btnRowGroup = optionDialog.add("group");
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        btnRowGroup.add("button", undefined, "OK", { name: "ok" });

        if (optionDialog.show() !== 1) return null;

        var includeItems = {};
        for (var i = 0; i < INFO_ITEMS.length; i++) {
            includeItems[INFO_ITEMS[i].key] = infoItemsUI.checkboxes[INFO_ITEMS[i].key].value;
        }
        return {
            position: positionRadios.posBottom.value ? "bottom" : "right",
            displayMode: modeRadios.modeFull.value ? "full" : "compact",
            includeItems: includeItems
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 文字を編集中で、そのストーリーが1つのテキストだけなら、テキストオブジェクトの選択に切り替える
     * @returns {void}
     */
    function selectFrameOfEditedText() {
        if (app.selection.constructor.name !== "TextRange") return;

        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall");
            app.selection = [textFramesInStory[0]];
            /* 選択ツールに切り替え（失敗しても続行） / Switch to the Selection tool; ignore failures */
            try { app.selectTool("Adobe Select Tool"); } catch (e) {}
        }
    }

    /**
     * 選択したテキストごとにフォント情報のラベルを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        selectFrameOfEditedText();

        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0 || selectedItems.length >= MAX_SELECTION_COUNT) return;

        var labelOptions = showOptionDialog();
        if (!labelOptions) return;

        var infoLayer = getOrCreateLayer(doc, INFO_LAYER_NAME);
        var generatedItems = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var sourceFrame = selectedItems[i];
            if (sourceFrame.typename !== "TextFrame") continue;

            var sourceFont = getFirstAvailableFont(sourceFrame);
            if (!sourceFont) continue;

            var textInfo = readTextInfo(sourceFrame, sourceFont, labelOptions.displayMode);
            var sourceBounds = sourceFrame.geometricBounds;
            var infoFrame;

            if (labelOptions.displayMode === "compact") {
                infoFrame = createCompactInfoFrame(doc, buildCompactText(textInfo), sourceBounds, labelOptions.position);
            } else {
                infoFrame = createDetailedInfoFrame(doc, buildDetailedText(textInfo, labelOptions.includeItems), sourceBounds, labelOptions.position);
            }
            applyInfoFrameStyle(infoFrame, labelOptions.displayMode);

            infoFrame.move(infoLayer, ElementPlacement.PLACEATBEGINNING);
            generatedItems.push(infoFrame);
            sourceFrame.selected = false;
        }

        if (generatedItems.length > 0) {
            app.selection = generatedItems;
        }
        app.redraw(); /* 画面再描画 / Redraw the screen */
    }

    main();

})();
