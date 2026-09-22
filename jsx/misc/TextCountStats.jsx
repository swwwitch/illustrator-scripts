#target illustrator
#targetengine "TextCountStatsSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストの文字数などを集計して表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStats.md

### Overview

Tallies the character counts and related statistics of the text in the document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStats.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextCountStats";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStats.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStats.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var LABEL_WIDTH = 110;                  /* 項目名の幅 / Row label width */
    var VALUE_WIDTH = 100;                  /* 値の幅 / Value width */
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネルの余白 / Panel margins */
    var PALETTE_OPACITY = 0.97;             /* パレットの不透明度 / Palette opacity */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "文字数カウント", en: "Text Count Stats" }
        },
        panel: {
            charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
            check: { ja: "チェック項目", en: "Check Items" },
            kinds: { ja: "種別", en: "Type" },
            other: { ja: "その他", en: "Other" }
        },
        fieldLabel: {
            chars: { ja: "文字", en: "Characters" },
            paras: { ja: "段落", en: "Paragraphs" },
            lines: { ja: "行", en: "Lines" },
            words: { ja: "英単語", en: "English Words" },
            fullwidth: { ja: "全角文字", en: "Fullwidth Chars" },
            hankakuKana: { ja: "半角カナ", en: "Half-width Kana" },
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            pathText: { ja: "パス上文字", en: "Type on a Path" },
            fonts: { ja: "使用フォント", en: "Fonts Used" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" }
        },
        status: {
            ready: { ja: "準備完了", en: "Ready" },
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
            wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
            selectedPrefix: { ja: "選択 ", en: "Selected " },
            selectedSuffix: { ja: " 件を集計", en: " object(s)" },
            timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
            busy: { ja: "処理中です", en: "Busy" },
            error: { ja: "エラー", en: "Error" }
        },
        tooltip: {
            refresh: { ja: "選択内容を再集計（⌘R / Enter）", en: "Recount selection (Cmd+R / Enter)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" },
            statValue: {
                ja: "左が選択範囲、右がドキュメント全体の値です",
                en: "Shows the selection on the left and the whole document on the right"
            },
            paras: { ja: "空の段落や空白だけの段落は数えません", en: "Empty paragraphs and paragraphs with only spaces are not counted" },
            lines: { ja: "自動で折り返された行も1行として数えます", en: "Lines created by automatic wrapping are counted too" },
            words: { ja: "半角英字だけの並びを1語として数えます", en: "Counts each run of ASCII letters as one word" },
            fullwidth: {
                ja: "全角の英数字・記号（！〜｠、￠〜￦）を数えます。漢字・かな・句読点は含みません",
                en: "Counts fullwidth letters, digits and symbols (U+FF01-FF60, U+FFE0-FFE6); kanji, kana and Japanese punctuation are not included"
            },
            fonts: {
                ja: "選択に関係なく、ドキュメント全体で使われているフォントの数を表示します",
                en: "Shows the number of fonts used in the whole document, regardless of the selection"
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /* ============================================================
       worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
       ------------------------------------------------------------
       注意 / Notes:
       - toString() で送るので JSDoc を付けない。// 行コメント禁止・/* *\/ のみ・
         各文は必ずセミコロンで終える
       - Sent via toString(): no JSDoc, no // comments; use block comments and
         always terminate statements with a semicolon
       ============================================================ */
    function wkCountTextStats() {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection) { currentSelection = []; }
        var fontSet = {};

        /* テキストフレームを種別・文字数などで集計し、フォント名を fontSet に集める / Tally text frames and collect font names */
        function tallyTextFrames(items) {
            var tally = { chars: 0, paras: 0, lines: 0, words: 0, fullwidth: 0, kana: 0, point: 0, area: 0, path: 0 };
            for (var i = 0; i < items.length; i++) {
                var frame = items[i];
                if (frame.typename !== "TextFrame") { continue; }
                if (frame.kind === TextType.POINTTEXT) { tally.point++; }
                else if (frame.kind === TextType.AREATEXT) { tally.area++; }
                else if (frame.kind === TextType.PATHTEXT) { tally.path++; }
                try {
                    var frameRange = frame.textRange || frame.textRanges[0];
                    fontSet[frameRange.characterAttributes.textFont.name] = true;
                } catch (e) {}
                try { tally.chars += frame.characters.length; } catch (e2) {}
                /* 末尾の空段落は読むと例外になることがある / Reading a trailing empty paragraph can throw */
                try {
                    var paragraphs = frame.paragraphs;
                    var filledParaCount = 0;
                    for (var p = 0; p < paragraphs.length; p++) {
                        if (paragraphs[p].contents.replace(/[\s　]/g, "").length > 0) { filledParaCount++; }
                    }
                    tally.paras += filledParaCount;
                } catch (e3) {}
                try { tally.lines += frame.lines.length; } catch (e4) {}
                try {
                    var frameContents = frame.contents;
                    if (typeof frameContents === "string") {
                        var wordMatches = frameContents.match(/\b[a-zA-Z]+\b/g); if (wordMatches) { tally.words += wordMatches.length; }
                        var fullwidthMatches = frameContents.match(/[！-｠￠-￦]/g); if (fullwidthMatches) { tally.fullwidth += fullwidthMatches.length; }
                        var kanaMatches = frameContents.match(/[･-ﾟ]/g); if (kanaMatches) { tally.kana += kanaMatches.length; }
                    }
                } catch (e5) {}
            }
            return tally;
        }

        var selectionTally = tallyTextFrames(currentSelection);
        var allTally = tallyTextFrames(doc.pageItems);

        var fontCount = 0;
        for (var fontName in fontSet) { if (fontSet.hasOwnProperty(fontName)) { fontCount++; } }

        var resultPairs = [
            "selCount=" + currentSelection.length,
            "charSel=" + selectionTally.chars,
            "charAll=" + allTally.chars,
            "paraSel=" + selectionTally.paras,
            "paraAll=" + allTally.paras,
            "lineSel=" + selectionTally.lines,
            "lineAll=" + allTally.lines,
            "wordSel=" + selectionTally.words,
            "wordAll=" + allTally.words,
            "fwSel=" + selectionTally.fullwidth,
            "fwAll=" + allTally.fullwidth,
            "kanaSel=" + selectionTally.kana,
            "kanaAll=" + allTally.kana,
            "pointSel=" + selectionTally.point,
            "pointAll=" + allTally.point,
            "areaSel=" + selectionTally.area,
            "areaAll=" + allTally.area,
            "pathSel=" + selectionTally.path,
            "pathAll=" + allTally.path,
            "fontCount=" + fontCount
        ];
        return "OK|" + resultPairs.join("|");
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [wkCountTextStats];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================
    var isBridgeBusy = false;

    /**
     * worker 関数を連結してメインエンジンで評価し、結果の文字列を返す
     * @param {string} callExpression - 評価させる呼び出し式（例: "wkCountTextStats()"）
     * @returns {string} worker の戻り値、または "ERR:BUSY" / "ERR:TIMEOUT" / "ERR:…"
     */
    function callMainEngine(callExpression) {
        /* 再入防止 / Re-entrancy guard */
        if (isBridgeBusy) { return "ERR:BUSY"; }
        isBridgeBusy = true;

        var resultHolder = { value: null };
        /* BridgeTalk の生成・送信 / Creating and sending BridgeTalk */
        try {
            /* worker 群を連結し、末尾に呼び出し式を付与 / Concatenate workers + call */
            var workerSource = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { workerSource += WORKER_FUNCS[i].toString(); }
            workerSource += callExpression + ";";

            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            /* encodeURIComponent で多バイト・改行・特殊文字の破損を回避 / Avoid corrupting multibyte text and newlines */
            bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\"));";
            bridgeTalk.onResult = function (response) {
                resultHolder.value = (response && response.body != null) ? String(response.body) : "";
            };
            bridgeTalk.onError = function (errorResponse) {
                resultHolder.value = "ERR:" + ((errorResponse && errorResponse.body) ? errorResponse.body : "bridge");
            };
            bridgeTalk.send(10);
        } catch (e) {
            resultHolder.value = "ERR:" + e;
        } finally {
            isBridgeBusy = false;
        }

        if (resultHolder.value === null) { return "ERR:TIMEOUT"; }
        return resultHolder.value;
    }

    /**
     * worker の戻り値（OK|key=value|...）を解析する
     * @param {string} response - worker の戻り値
     * @returns {Object|null} キーと値の対応表（形式が違うときは null）
     */
    function parseStats(response) {
        if (!response || response.indexOf("OK|") !== 0) return null;
        var statsMap = {};
        var pairs = response.substring(3).split("|");
        for (var i = 0; i < pairs.length; i++) {
            var keyValue = pairs[i].split("=");
            if (keyValue.length === 2) { statsMap[keyValue[0]] = keyValue[1]; }
        }
        return statsMap;
    }

    // =========================================
    // パレット構築 / Build palette
    // =========================================

    /**
     * 集計項目をまとめるパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titlePath - パネルタイトルの LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addStatPanel(parentGroup, titlePath) {
        var statPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        statPanel.orientation = "column";
        statPanel.alignChildren = ["fill", "top"];
        statPanel.margins = PANEL_MARGINS;
        return statPanel;
    }

    /**
     * 項目名と値の行を追加し、値の statictext を返す
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel のキー（LABELS.tooltip に同じキーがあれば項目名の tooltip にする）
     * @returns {StaticText} 値を表示する statictext
     */
    function addStatRow(parentPanel, labelKey) {
        var statRow = parentPanel.add("group");
        statRow.orientation = "row";
        statRow.alignChildren = ["left", "center"];

        var rowLabel = statRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        if (LABELS.tooltip[labelKey]) rowLabel.helpTip = getLabel("tooltip." + labelKey);

        var valueText = statRow.add("statictext", undefined, "-");
        valueText.preferredSize.width = VALUE_WIDTH;
        valueText.justify = "left";
        return valueText;
    }

    /**
     * 集計結果を値の欄に書き込む
     * @param {Object} valueFields - 項目キーと値の statictext の対応表
     * @param {Object} stats - parseStats() の結果
     * @returns {void}
     */
    function showStats(valueFields, stats) {
        valueFields.chars.text = stats.charSel + " / " + stats.charAll;
        valueFields.paras.text = stats.paraSel + " / " + stats.paraAll;
        valueFields.lines.text = stats.lineSel + " / " + stats.lineAll;
        valueFields.words.text = stats.wordSel + " / " + stats.wordAll;
        valueFields.fullwidth.text = stats.fwSel + " / " + stats.fwAll;
        valueFields.hankakuKana.text = stats.kanaSel + " / " + stats.kanaAll;
        valueFields.pointText.text = stats.pointSel + " / " + stats.pointAll;
        valueFields.areaText.text = stats.areaSel + " / " + stats.areaAll;
        valueFields.pathText.text = stats.pathSel + " / " + stats.pathAll;
        valueFields.fonts.text = stats.fontCount;
    }

    /**
     * 集計パレットを組み立てる
     * @returns {Window} 組み立てたパレット
     */
    function buildPalette() {
        var statsPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        statsPalette.orientation = "column";
        statsPalette.alignChildren = ["fill", "top"];

        var panelColumnGroup = statsPalette.add("group");
        panelColumnGroup.orientation = "column";
        panelColumnGroup.alignChildren = ["fill", "top"];

        var charParaPanel = addStatPanel(panelColumnGroup, "panel.charPara");
        var checkPanel = addStatPanel(panelColumnGroup, "panel.check");
        var kindsPanel = addStatPanel(panelColumnGroup, "panel.kinds");
        var otherPanel = addStatPanel(panelColumnGroup, "panel.other");

        /* 値の statictext 参照を保持 / Keep references to value fields */
        var valueFields = {
            chars: addStatRow(charParaPanel, "chars"),
            paras: addStatRow(charParaPanel, "paras"),
            lines: addStatRow(charParaPanel, "lines"),
            words: addStatRow(charParaPanel, "words"),
            fullwidth: addStatRow(checkPanel, "fullwidth"),
            hankakuKana: addStatRow(checkPanel, "hankakuKana"),
            pointText: addStatRow(kindsPanel, "pointText"),
            areaText: addStatRow(kindsPanel, "areaText"),
            pathText: addStatRow(kindsPanel, "pathText"),
            fonts: addStatRow(otherPanel, "fonts")
        };
        /* 「選択 / 全体」の並びを説明（フォント数は全体のみ） / Explain the "selection / all" pair (fonts show the total only) */
        for (var fieldKey in valueFields) {
            if (fieldKey !== "fonts") valueFields[fieldKey].helpTip = getLabel("tooltip.statValue");
        }

        /* ステータス表示 / Status line */
        var statusText = statsPalette.add("statictext", undefined, getLabel("status.ready"));
        statusText.alignment = ["fill", "bottom"];

        /**
         * ステータス行の表示を差し替える
         * @param {string} message - 表示する文字列
         * @returns {void}
         */
        function setStatus(message) {
            statusText.text = message;
        }

        /**
         * メインエンジンで集計し直して表示を更新する
         * @returns {void}
         */
        function refreshStats() {
            setStatus(getLabel("status.busy"));
            var response = callMainEngine("wkCountTextStats()");

            if (response === "ERR:BUSY") { setStatus(getLabel("status.busy")); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel("status.timeout")); return; }
            if (response === "NODOC") { setStatus(getLabel("status.noDoc")); return; }
            if (response.indexOf("ERR:") === 0) { setStatus(getLabel("status.error") + ": " + response.substring(4)); return; }

            var stats = parseStats(response);
            if (!stats) { setStatus(getLabel("status.error")); return; }

            showStats(valueFields, stats);

            var selectedCount = parseInt(stats.selCount, 10) || 0;
            if (selectedCount > 0) {
                setStatus(getLabel("status.selectedPrefix") + selectedCount + getLabel("status.selectedSuffix"));
            } else {
                setStatus(getLabel("status.wholeDoc"));
            }
        }

        /* ボタン（更新のみ。閉じるは × / Esc に任せる） / Button (refresh only) */
        var btnRowGroup = statsPalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.alignChildren = ["right", "center"];

        var btnRefresh = btnRowGroup.add("button", undefined, getLabel("button.refresh"));
        btnRefresh.helpTip = getLabel("tooltip.refresh") + "\n" + getLabel("tooltip.esc");
        /* onClick 連結（addEventListener('click') は不発の環境がある） / addEventListener('click') does not fire in some environments */
        btnRefresh.onClick = refreshStats;

        /* キー操作：Esc で閉じる、⌘R / Enter で更新 / Keys: Esc close, Cmd+R / Enter refresh */
        statsPalette.addEventListener("keydown", function (keyEvent) {
            var keyName = (keyEvent && keyEvent.keyName) ? String(keyEvent.keyName).toUpperCase() : "";
            if (keyName === "ESCAPE") {
                statsPalette.close();
            } else if (keyName === "R" || keyName === "ENTER" || keyName === "RETURN") {
                refreshStats();
            }
        });

        /* opacity に対応しない環境がある / Some environments do not support opacity */
        try { statsPalette.opacity = PALETTE_OPACITY; } catch (e) {}

        /* 表示直後に一度集計 / Count once right after showing */
        statsPalette.onShow = function () {
            refreshStats();
        };

        return statsPalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * パレットを開く（開いているものがあれば閉じてから開き直す）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる（前回の実行で無効になっていることがある）
           Prevent duplicates: close the existing palette (it may already be invalid) */
        if ($.global.__TextCountStatsPalette) {
            try { $.global.__TextCountStatsPalette.close(); } catch (e) {}
            $.global.__TextCountStatsPalette = null;
        }

        var statsPalette = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__TextCountStatsPalette = statsPalette;
        statsPalette.onClose = function () {
            $.global.__TextCountStatsPalette = null;
        };

        statsPalette.center();
        statsPalette.show();
    }

    showPalette();

})();
