#target illustrator
#targetengine "TextCountStatsSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

TextCountStats.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStats.jsx

### 概要：

- Illustrator の選択テキストや全体の文字情報を統計的に可視化
- 文字数、段落数、行数、英単語数、全角文字数、半角カナ、種別（ポイント／エリア／パス上）、フォント数などをカウント
- 常駐パレット化：開いたまま選択を切り替え、［更新］で再集計

### 主な機能：

- 選択したテキストオブジェクトの各種統計をパレット上に一覧表示
- 選択がない場合はドキュメント全体のテキストを対象に集計
- DOM 集計はメインエンジンへ BridgeTalk 委譲（常駐パレットの DOM 切断を回避）
- UI は日本語／英語対応

### 処理の流れ：

1. 常駐パレットを表示（多重起動は自動で閉じてから再表示）
2. ［更新］押下（または表示直後）にメインエンジンへ集計を委譲
3. 戻り値（マーカー方式）を解析し、各パネルの値を更新

### 更新履歴：

- v1.0 (20250806) : 初期バージョン
- v1.1 (20260702) : 常駐パレット化（#targetengine ＋ BridgeTalk 委譲）、更新ボタン・ステータス表示・ローカライズ整理

---

### Script Name:

TextCountStats.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStats.jsx

### Description:

- Visualize statistics of selected or all text objects in Illustrator
- Count characters, paragraphs, lines, English words, full-width characters, half-width kana, type (point/area/path), and fonts
- Persistent palette: keep it open, change selection, press Refresh to recount

### Main Features:

- Shows a summary of various counts in a palette
- Automatically switches between selected objects and all content
- Delegates DOM counting to the main engine via BridgeTalk (avoids palette DOM disconnection)
- UI supports Japanese and English

### Flow:

1. Show a persistent palette (existing one is closed first to prevent duplicates)
2. On Refresh (or right after showing), delegate counting to the main engine
3. Parse the marker-based result and update each panel value

### Change Log:

- v1.0 (20250806): Initial version
- v1.1 (20260702): Palette conversion (#targetengine + BridgeTalk delegation), refresh button, status line, localization cleanup

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextCountStats";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* ============================================================
   言語判定・ローカライズ / Language & localization
   ============================================================ */
function getCurrentLang() {
    /* 言語判定 / Determine language */
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

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

/* L(): ドットパス参照（null 耐性） / Dot-path lookup with null tolerance */
function L(path) {
    var parts = String(path).split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return path;
        node = node[parts[i]];
    }
    if (node == null) return path;
    if (typeof node === "string") return node;
    if (typeof node === "object" && node[lang] != null) return node[lang];
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
    var sel = doc.selection;
    if (!sel) { sel = []; }
    var selCount = sel.length;

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

    for (var si = 0; si < sel.length; si++) {
        if (sel[si].typename === "TextFrame") {
            if (sel[si].kind === TextType.POINTTEXT) { pointTextSel++; }
            else if (sel[si].kind === TextType.AREATEXT) { areaTextSel++; }
            else if (sel[si].kind === TextType.PATHTEXT) { pathTextSel++; }
            try {
                var ran = sel[si].textRange || sel[si].textRanges[0];
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

    for (var sj = 0; sj < sel.length; sj++) {
        if (sel[sj].typename === "TextFrame") {
            try { totalCharSel += sel[sj].characters.length; } catch (e3) {}
            try {
                var paras = sel[sj].paragraphs;
                var vp = 0;
                for (var p = 0; p < paras.length; p++) {
                    var c = paras[p].contents;
                    if (c.replace(/[\s　]/g, "").length > 0) { vp++; }
                }
                paraCountSel += vp;
            } catch (e4) {}
            try { lineCountSel += sel[sj].lines.length; } catch (e5) {}
            try {
                var cont = sel[sj].contents;
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
    var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    var columnGroup = win.add("group");
    columnGroup.orientation = "column";
    columnGroup.alignChildren = ["fill", "top"];

    /* 1. 文字・段落 / Characters & Paragraphs */
    var panelCharPara = columnGroup.add("panel", undefined, L('panel.charPara'));
    panelCharPara.orientation = "column";
    panelCharPara.alignChildren = ["fill", "top"];
    panelCharPara.margins = PANEL_MARGINS;

    /* 2. チェック項目 / Check items */
    var panelCheck = columnGroup.add("panel", undefined, L('panel.check'));
    panelCheck.orientation = "column";
    panelCheck.alignChildren = ["fill", "top"];
    panelCheck.margins = PANEL_MARGINS;

    /* 3. 種別 / Type */
    var panelKinds = columnGroup.add("panel", undefined, L('panel.kinds'));
    panelKinds.orientation = "column";
    panelKinds.alignChildren = ["fill", "top"];
    panelKinds.margins = PANEL_MARGINS;

    /* 4. その他 / Other */
    var panelOther = columnGroup.add("panel", undefined, L('panel.other'));
    panelOther.orientation = "column";
    panelOther.alignChildren = ["fill", "top"];
    panelOther.margins = PANEL_MARGINS;

    /* 値の statictext 参照を保持 / Keep references to value fields */
    var values = {
        chars: addRow(panelCharPara, L('row.chars')),
        paras: addRow(panelCharPara, L('row.paras')),
        lines: addRow(panelCharPara, L('row.lines')),
        words: addRow(panelCharPara, L('row.words')),
        fullwidth: addRow(panelCheck, L('row.fullwidth')),
        hankakuKana: addRow(panelCheck, L('row.hankakuKana')),
        pointText: addRow(panelKinds, L('row.pointText')),
        areaText: addRow(panelKinds, L('row.areaText')),
        pathText: addRow(panelKinds, L('row.pathText')),
        fonts: addRow(panelOther, L('row.fonts'))
    };

    /* ステータス表示 / Status line */
    var statusText = win.add("statictext", undefined, L('status.ready'));
    statusText.alignment = ["fill", "bottom"];

    function setStatus(msg) {
        try { statusText.text = msg; } catch (e) {}
    }

    /* 再集計 / Recount */
    function refresh() {
        setStatus(L('status.busy'));
        var resp = callMainEngine("wkCountTextStats()");

        if (resp === "ERR:BUSY") { setStatus(L('status.busy')); return; }
        if (resp === null || resp === "ERR:TIMEOUT") { setStatus(L('status.timeout')); return; }
        if (resp === "NODOC") { setStatus(L('status.noDoc')); return; }
        if (resp.indexOf("ERR:") === 0) { setStatus(L('status.error') + ": " + resp.substring(4)); return; }

        var m = parseStats(resp);
        if (!m) { setStatus(L('status.error')); return; }

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
            setStatus(L('status.selectedPrefix') + selN + L('status.selectedSuffix'));
        } else {
            setStatus(L('status.wholeDoc'));
        }
    }

    /* ボタン（更新のみ。閉じるは × / Esc に任せる） / Button (refresh only) */
    var btnRow = win.add("group");
    btnRow.orientation = "row";
    btnRow.alignment = ["fill", "bottom"];
    btnRow.alignChildren = ["right", "center"];

    var btnRefresh = btnRow.add("button", undefined, L('button.refresh'));
    btnRefresh.helpTip = L('hint.refresh') + "\n" + L('hint.esc');
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
        try { $.global.__TextCountStatsPalette = null; } catch (e) {}
    };

    win.center();
    win.show();
}

showPalette();
