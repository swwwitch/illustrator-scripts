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
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStats.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStats.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        /* 言語判定 / Determine language */
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
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
        row: {
            chars: { ja: "文字：", en: "Characters:" },
            paras: { ja: "段落：", en: "Paragraphs:" },
            lines: { ja: "行：", en: "Lines:" },
            words: { ja: "英単語：", en: "English Words:" },
            fullwidth: { ja: "全角文字：", en: "Fullwidth Chars:" },
            hankakuKana: { ja: "半角カナ：", en: "Half-width Kana:" },
            pointText: { ja: "ポイント文字：", en: "Point Type:" },
            areaText: { ja: "エリア内文字：", en: "Area Type:" },
            pathText: { ja: "パス上文字：", en: "Type on a Path:" },
            fonts: { ja: "使用フォント：", en: "Fonts Used:" }
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
        hint: {
            refresh: { ja: "選択内容を再集計（⌘R / Enter）", en: "Recount selection (Cmd+R / Enter)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" }
        }
    };

    /* getLabel(): ドットパス参照（null 耐性） / Dot-path lookup with null tolerance */
    function getLabel(path) {
        var parts = String(path).split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return path;
            node = node[parts[i]];
        }
        if (node == null) return path;
        if (typeof node === "string") return node;
        if (typeof node === "object" && node[uiLang] != null) return node[uiLang];
        return path;
    }

    /* ============================================================
       定数 / Constants
       ============================================================ */
    var LABEL_WIDTH = 110;
    var VALUE_WIDTH = 100;
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PALETTE_OPACITY = 0.97;

    /* ============================================================
       worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
       ------------------------------------------------------------
       注意 / Notes:
       - toString() は改行を全削除するため、// 行コメント禁止・/* *\/ のみ・
         各文は必ずセミコロンで終える
       - toString strips newlines, so no // comments; use block comments and
         always terminate statements with a semicolon
       ============================================================ */
    function wkCountTextStats() {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection) { currentSelection = []; }
        var selCount = currentSelection.length;

        var allCount = 0;
        function countAll(items) {
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                allCount++;
                if (it.typename === "GroupItem") { countAll(it.pageItems); }
                else if (it.typename === "CompoundPathItem") { countAll(it.pathItems); }
            }
        }
        countAll(doc.pageItems);

        var pointTextSel = 0, areaTextSel = 0, pathTextSel = 0;
        var pointTextAll = 0, areaTextAll = 0, pathTextAll = 0;
        var fontSet = {};

        for (var si = 0; si < currentSelection.length; si++) {
            if (currentSelection[si].typename === "TextFrame") {
                if (currentSelection[si].kind === TextType.POINTTEXT) { pointTextSel++; }
                else if (currentSelection[si].kind === TextType.AREATEXT) { areaTextSel++; }
                else if (currentSelection[si].kind === TextType.PATHTEXT) { pathTextSel++; }
                try {
                    var ran = currentSelection[si].textRange || currentSelection[si].textRanges[0];
                    var fnt = ran.characterAttributes.textFont.name;
                    fontSet[fnt] = true;
                } catch (e) {}
            }
        }

        var allItems = doc.pageItems;
        for (var ai = 0; ai < allItems.length; ai++) {
            var obj = allItems[ai];
            if (obj.typename === "TextFrame") {
                if (obj.kind === TextType.POINTTEXT) { pointTextAll++; }
                else if (obj.kind === TextType.AREATEXT) { areaTextAll++; }
                else if (obj.kind === TextType.PATHTEXT) { pathTextAll++; }
                try {
                    var ran2 = obj.textRange || obj.textRanges[0];
                    var fnt2 = ran2.characterAttributes.textFont.name;
                    fontSet[fnt2] = true;
                } catch (e2) {}
            }
        }

        var totalCharSel = 0, totalCharAll = 0;
        var paraCountSel = 0, paraCountAll = 0;
        var wordCountSel = 0, wordCountAll = 0;
        var fullwidthSel = 0, fullwidthAll = 0;
        var kanaSel = 0, kanaAll = 0;
        var lineCountSel = 0, lineCountAll = 0;

        for (var sj = 0; sj < currentSelection.length; sj++) {
            if (currentSelection[sj].typename === "TextFrame") {
                try { totalCharSel += currentSelection[sj].characters.length; } catch (e3) {}
                try {
                    var paras = currentSelection[sj].paragraphs;
                    var vp = 0;
                    for (var p = 0; p < paras.length; p++) {
                        var c = paras[p].contents;
                        if (c.replace(/[\s　]/g, "").length > 0) { vp++; }
                    }
                    paraCountSel += vp;
                } catch (e4) {}
                try { lineCountSel += currentSelection[sj].lines.length; } catch (e5) {}
                try {
                    var cont = currentSelection[sj].contents;
                    if (typeof cont === "string") {
                        var mw = cont.match(/\b[a-zA-Z]+\b/g); if (mw) { wordCountSel += mw.length; }
                        var mf = cont.match(/[！-｠￠-￦]/g); if (mf) { fullwidthSel += mf.length; }
                        var mk = cont.match(/[･-ﾟ]/g); if (mk) { kanaSel += mk.length; }
                    }
                } catch (e6) {}
            }
        }

        for (var aj = 0; aj < allItems.length; aj++) {
            var obj2 = allItems[aj];
            if (obj2.typename === "TextFrame") {
                try { totalCharAll += obj2.characters.length; } catch (e7) {}
                try {
                    var paras2 = obj2.paragraphs;
                    var vp2 = 0;
                    for (var p2 = 0; p2 < paras2.length; p2++) {
                        var c2 = paras2[p2].contents;
                        if (c2.replace(/[\s　]/g, "").length > 0) { vp2++; }
                    }
                    paraCountAll += vp2;
                } catch (e8) {}
                try { lineCountAll += obj2.lines.length; } catch (e9) {}
                try {
                    var cont2 = obj2.contents;
                    if (typeof cont2 === "string") {
                        var mw2 = cont2.match(/\b[a-zA-Z]+\b/g); if (mw2) { wordCountAll += mw2.length; }
                        var mf2 = cont2.match(/[！-｠￠-￦]/g); if (mf2) { fullwidthAll += mf2.length; }
                        var mk2 = cont2.match(/[･-ﾟ]/g); if (mk2) { kanaAll += mk2.length; }
                    }
                } catch (e10) {}
            }
        }

        var fontCount = 0;
        for (var key in fontSet) { if (fontSet.hasOwnProperty(key)) { fontCount++; } }

        var out = [];
        out.push("selCount=" + selCount);
        out.push("allCount=" + allCount);
        out.push("charSel=" + totalCharSel);
        out.push("charAll=" + totalCharAll);
        out.push("paraSel=" + paraCountSel);
        out.push("paraAll=" + paraCountAll);
        out.push("lineSel=" + lineCountSel);
        out.push("lineAll=" + lineCountAll);
        out.push("wordSel=" + wordCountSel);
        out.push("wordAll=" + wordCountAll);
        out.push("fwSel=" + fullwidthSel);
        out.push("fwAll=" + fullwidthAll);
        out.push("kanaSel=" + kanaSel);
        out.push("kanaAll=" + kanaAll);
        out.push("pointSel=" + pointTextSel);
        out.push("pointAll=" + pointTextAll);
        out.push("areaSel=" + areaTextSel);
        out.push("areaAll=" + areaTextAll);
        out.push("pathSel=" + pathTextSel);
        out.push("pathAll=" + pathTextAll);
        out.push("fontCount=" + fontCount);
        return "OK|" + out.join("|");
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [wkCountTextStats];

    /* ============================================================
       BridgeTalk 委譲 / Delegation to the main engine
       ============================================================ */
    var isBusy = false;

    function callMainEngine(callExpr) {
        /* 再入防止 / Re-entrancy guard */
        if (isBusy) { return "ERR:BUSY"; }
        isBusy = true;

        var holder = { value: null };
        try {
            /* worker 群を連結し、末尾に呼び出し式を付与 / Concatenate workers + call */
            var src = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { src += WORKER_FUNCS[i].toString(); }
            src += callExpr + ";";

            var bt = new BridgeTalk();
            bt.target = "illustrator";
            /* encodeURIComponent で多バイト・改行・特殊文字の破損を回避 */
            bt.body = "eval(decodeURIComponent(\"" + encodeURIComponent(src) + "\"));";
            bt.onResult = function (res) {
                holder.value = (res && res.body != null) ? String(res.body) : "";
            };
            bt.onError = function (err) {
                holder.value = "ERR:" + ((err && err.body) ? err.body : "bridge");
            };
            bt.send(10);
        } catch (e) {
            holder.value = "ERR:" + e;
        } finally {
            isBusy = false;
        }

        if (holder.value === null) { return "ERR:TIMEOUT"; }
        return holder.value;
    }

    /* 戻り値（OK|key=value|...）を解析 / Parse "OK|key=value|..." */
    function parseStats(resp) {
        if (!resp || resp.indexOf("OK|") !== 0) return null;
        var map = {};
        var parts = resp.substring(3).split("|");
        for (var i = 0; i < parts.length; i++) {
            var kv = parts[i].split("=");
            if (kv.length === 2) { map[kv[0]] = kv[1]; }
        }
        return map;
    }

    /* ============================================================
       パレット構築 / Build palette
       ============================================================ */
    function addRow(panel, labelText) {
        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];

        var label = row.add("statictext", undefined, labelText);
        label.preferredSize.width = LABEL_WIDTH;
        label.justify = "right";

        var value = row.add("statictext", undefined, "-");
        value.preferredSize.width = VALUE_WIDTH;
        value.justify = "left";
        return value;
    }

    function buildPalette() {
        var win = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];

        var columnGroup = win.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];

        /* 1. 文字・段落 / Characters & Paragraphs */
        var panelCharPara = columnGroup.add("panel", undefined, getLabel('panel.charPara'));
        panelCharPara.orientation = "column";
        panelCharPara.alignChildren = ["fill", "top"];
        panelCharPara.margins = PANEL_MARGINS;

        /* 2. チェック項目 / Check items */
        var panelCheck = columnGroup.add("panel", undefined, getLabel('panel.check'));
        panelCheck.orientation = "column";
        panelCheck.alignChildren = ["fill", "top"];
        panelCheck.margins = PANEL_MARGINS;

        /* 3. 種別 / Type */
        var panelKinds = columnGroup.add("panel", undefined, getLabel('panel.kinds'));
        panelKinds.orientation = "column";
        panelKinds.alignChildren = ["fill", "top"];
        panelKinds.margins = PANEL_MARGINS;

        /* 4. その他 / Other */
        var panelOther = columnGroup.add("panel", undefined, getLabel('panel.other'));
        panelOther.orientation = "column";
        panelOther.alignChildren = ["fill", "top"];
        panelOther.margins = PANEL_MARGINS;

        /* 値の statictext 参照を保持 / Keep references to value fields */
        var values = {
            chars: addRow(panelCharPara, getLabel('row.chars')),
            paras: addRow(panelCharPara, getLabel('row.paras')),
            lines: addRow(panelCharPara, getLabel('row.lines')),
            words: addRow(panelCharPara, getLabel('row.words')),
            fullwidth: addRow(panelCheck, getLabel('row.fullwidth')),
            hankakuKana: addRow(panelCheck, getLabel('row.hankakuKana')),
            pointText: addRow(panelKinds, getLabel('row.pointText')),
            areaText: addRow(panelKinds, getLabel('row.areaText')),
            pathText: addRow(panelKinds, getLabel('row.pathText')),
            fonts: addRow(panelOther, getLabel('row.fonts'))
        };

        /* ステータス表示 / Status line */
        var statusText = win.add("statictext", undefined, getLabel('status.ready'));
        statusText.alignment = ["fill", "bottom"];

        function setStatus(msg) {
            statusText.text = msg;
        }

        /* 再集計 / Recount */
        function refresh() {
            setStatus(getLabel('status.busy'));
            var resp = callMainEngine("wkCountTextStats()");

            if (resp === "ERR:BUSY") { setStatus(getLabel('status.busy')); return; }
            if (resp === null || resp === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
            if (resp === "NODOC") { setStatus(getLabel('status.noDoc')); return; }
            if (resp.indexOf("ERR:") === 0) { setStatus(getLabel('status.error') + ": " + resp.substring(4)); return; }

            var m = parseStats(resp);
            if (!m) { setStatus(getLabel('status.error')); return; }

            values.chars.text = m.charSel + " / " + m.charAll;
            values.paras.text = m.paraSel + " / " + m.paraAll;
            values.lines.text = m.lineSel + " / " + m.lineAll;
            values.words.text = m.wordSel + " / " + m.wordAll;
            values.fullwidth.text = m.fwSel + " / " + m.fwAll;
            values.hankakuKana.text = m.kanaSel + " / " + m.kanaAll;
            values.pointText.text = m.pointSel + " / " + m.pointAll;
            values.areaText.text = m.areaSel + " / " + m.areaAll;
            values.pathText.text = m.pathSel + " / " + m.pathAll;
            values.fonts.text = m.fontCount;

            var selN = parseInt(m.selCount, 10) || 0;
            if (selN > 0) {
                setStatus(getLabel('status.selectedPrefix') + selN + getLabel('status.selectedSuffix'));
            } else {
                setStatus(getLabel('status.wholeDoc'));
            }
        }

        /* ボタン（更新のみ。閉じるは × / Esc に任せる） / Button (refresh only) */
        var btnRow = win.add("group");
        btnRow.orientation = "row";
        btnRow.alignment = ["fill", "bottom"];
        btnRow.alignChildren = ["right", "center"];

        var btnRefresh = btnRow.add("button", undefined, getLabel('button.refresh'));
        btnRefresh.helpTip = getLabel('hint.refresh') + "\n" + getLabel('hint.esc');
        /* onClick 連結（addEventListener('click') は不発の環境がある） */
        btnRefresh.onClick = refresh;

        /* キー操作：Esc で閉じる、⌘R / Enter で更新 / Keys: Esc close, Cmd+R / Enter refresh */
        win.addEventListener("keydown", function (ev) {
            var k = "";
            try { k = ev && ev.keyName ? String(ev.keyName).toUpperCase() : ""; } catch (e) { k = ""; }
            if (k === "ESCAPE") {
                try { win.close(); } catch (e1) {}
            } else if (k === "R" || k === "ENTER" || k === "RETURN") {
                refresh();
            }
        });

        try { win.opacity = PALETTE_OPACITY; } catch (e) {}

        /* 表示直後に一度集計 / Count once right after showing */
        win.onShow = function () {
            refresh();
        };

        return win;
    }

    /* ============================================================
       起動 / Entry point
       ============================================================ */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる / Prevent duplicates */
        if ($.global.__TextCountStatsPalette) {
            try { $.global.__TextCountStatsPalette.close(); } catch (e) {}
            $.global.__TextCountStatsPalette = null;
        }

        var win = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__TextCountStatsPalette = win;
        win.onClose = function () {
            $.global.__TextCountStatsPalette = null;
        };

        win.center();
        win.show();
    }

    showPalette();

})();
