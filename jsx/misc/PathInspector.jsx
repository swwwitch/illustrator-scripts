#target illustrator
#targetengine "PathInspectorSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中またはドキュメント全体のパス統計を集計し、常駐パレットで表示します。

詳細は README を参照してください。

### Overview

Counts path statistics for the selection, or for the whole document, and shows them in a persistent palette.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathInspector";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspector.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspector.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 現在の言語コードを返す
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
var LABELS = {
    dialog: {
        title: { ja: "パスのカウント", en: "Path Count" }
    },
    report: {
        title: { ja: "Path Inspector Report", en: "Path Inspector Report" },
        document: { ja: "ドキュメント:", en: "Document:" },
        date: { ja: "日付:", en: "Date:" },
        valueNote: {
            ja: "※ 値は『選択 / 全体』の形式です",
            en: "Note: values are formatted as 'Selection / All'"
        }
    },
    section: {
        paths: { ja: "パス", en: "Paths" }
    },
    panel: {
        path: { ja: "パス", en: "Paths" }
    },
    row: {
        pathCount: { ja: "パス：", en: "Paths:" },
        openPath: { ja: "オープンパス：", en: "Open Paths:" },
        closedPath: { ja: "クローズパス：", en: "Closed Paths:" },
        anchors: { ja: "アンカーポイント：", en: "Anchor Points:" },
        handles: { ja: "ハンドル：", en: "Handles:" },
        compoundPath: { ja: "複合パス：", en: "Compound Paths:" },
        compoundShape: { ja: "複合シェイプ：", en: "Compound Shapes:" }
    },
    button: {
        refresh: { ja: "更新", en: "Refresh" },
        exportPreset: { ja: "書き出し", en: "Export" }
    },
    status: {
        ready: { ja: "準備完了", en: "Ready" },
        noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
        wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
        selectedPrefix: { ja: "選択 ", en: "Selected " },
        selectedSuffix: { ja: " 件を集計", en: " object(s)" },
        timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
        busy: { ja: "処理中です", en: "Busy" },
        error: { ja: "エラー", en: "Error" },
        exportedPrefix: { ja: "書き出しました: ", en: "Exported: " },
        exportFailOpen: { ja: "ファイルを開けませんでした", en: "Failed to open the file" }
    },
    hint: {
        refresh: { ja: "選択内容を再集計（⌘R）", en: "Recount selection (Cmd+R)" },
        esc: { ja: "Esc で閉じる", en: "Press Esc to close" }
    }
};

/**
 * ドットパスでラベルを参照する（見つからない場合はパス文字列を返す）
 * @param {string} path ラベルのドットパス（例: "row.pathCount"）
 * @returns {string} ローカライズ済み文字列
 */
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

/**
 * 書き出し用にラベル末尾のコロンを半角へ正規化する
 * @param {string} path ラベルのドットパス
 * @returns {string} 正規化済み文字列
 */
function LX(path) {
    return L(path).replace(/[：:]\s*$/, ":");
}

(function () {

    /* ============================================================
       定数 / Constants
       ============================================================ */
    var LABEL_WIDTH = 130;
    var VALUE_WIDTH = 90;
    var PANEL_MARGINS = [15, 20, 0, 10];
    var PALETTE_OPACITY = 0.97;

    /* ============================================================
       worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
       ------------------------------------------------------------
       注意 / Notes:
       - toString() は改行を全削除するため、// 行コメント禁止・/* *\/ のみ・
         各文は必ずセミコロンで終える
       ============================================================ */
    function wkIsGuidePath(pi) {
        try { return (pi && pi.typename === "PathItem" && pi.guides === true); } catch (e) { return false; }
    }

    function wkCountHandles(pi) {
        var c = 0;
        try {
            var pts = pi.pathPoints;
            for (var i = 0; i < pts.length; i++) {
                var p = pts[i];
                var a = p.anchor;
                var l = p.leftDirection;
                var rr = p.rightDirection;
                if (l[0] !== a[0] || l[1] !== a[1]) { c++; }
                if (rr[0] !== a[0] || rr[1] !== a[1]) { c++; }
            }
        } catch (e) {}
        return c;
    }

    function wkCountPathStats(it, stats) {
        if (it.typename === "GroupItem") {
            for (var gi = 0; gi < it.pageItems.length; gi++) { wkCountPathStats(it.pageItems[gi], stats); }
        } else if (it.typename === "PathItem") {
            if (!wkIsGuidePath(it)) {
                stats.pathCount++;
                stats.anchorCount += it.pathPoints.length;
                stats.handleCount += wkCountHandles(it);
                if (it.closed) { stats.closedPath++; } else { stats.openPath++; }
            }
        } else if (it.typename === "CompoundPathItem") {
            for (var ci = 0; ci < it.pathItems.length; ci++) {
                if (wkIsGuidePath(it.pathItems[ci])) { continue; }
                stats.pathCount++;
                stats.anchorCount += it.pathItems[ci].pathPoints.length;
                stats.handleCount += wkCountHandles(it.pathItems[ci]);
                if (it.pathItems[ci].closed) { stats.closedPath++; } else { stats.openPath++; }
            }
        }
    }

    function wkCollect() {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!sel) { sel = []; }
        var selCount = sel.length;

        var allItems = doc.pageItems;

        var cpathSel = 0, cpathAll = 0, cshapeSel = 0, cshapeAll = 0;

        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "CompoundPathItem") { cpathSel++; }
            if (sel[i].typename === "PluginItem") {
                try { if (sel[i].name && sel[i].name.indexOf("Compound Shape") !== -1) { cshapeSel++; } } catch (e) {}
            }
        }

        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "CompoundPathItem") { cpathAll++; }
            if (obj.typename === "PluginItem") {
                try { if (obj.name && obj.name.indexOf("Compound Shape") !== -1) { cshapeAll++; } } catch (e2) {}
            }
        }

        var pathStatsSel = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
        for (var i2 = 0; i2 < sel.length; i2++) { wkCountPathStats(sel[i2], pathStatsSel); }

        var pathStatsAll = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
        for (var k2 = 0; k2 < allItems.length; k2++) { wkCountPathStats(allItems[k2], pathStatsAll); }

        var docName = "";
        try { docName = doc.name; } catch (e3) { docName = ""; }

        var out = [];
        out.push("selCount=" + selCount);
        out.push("docName=" + encodeURIComponent(docName));
        out.push("pathCountSel=" + pathStatsSel.pathCount);
        out.push("pathCountAll=" + pathStatsAll.pathCount);
        out.push("openSel=" + pathStatsSel.openPath);
        out.push("openAll=" + pathStatsAll.openPath);
        out.push("closedSel=" + pathStatsSel.closedPath);
        out.push("closedAll=" + pathStatsAll.closedPath);
        out.push("anchorSel=" + pathStatsSel.anchorCount);
        out.push("anchorAll=" + pathStatsAll.anchorCount);
        out.push("handleSel=" + pathStatsSel.handleCount);
        out.push("handleAll=" + pathStatsAll.handleCount);
        out.push("cpathSel=" + cpathSel);
        out.push("cpathAll=" + cpathAll);
        out.push("cshapeSel=" + cshapeSel);
        out.push("cshapeAll=" + cshapeAll);

        return "OK|" + out.join("|");
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [
        wkIsGuidePath,
        wkCountHandles,
        wkCountPathStats,
        wkCollect
    ];

    /* ============================================================
       BridgeTalk 委譲 / Delegation to the main engine
       ============================================================ */
    var isBusy = false;

    /**
     * worker 関数群をメインエンジンへ送り、指定した式を評価して結果を得る
     * @param {string} callExpr メインエンジンで評価する式（例: "wkCollect()"）
     * @returns {string} 戻り値文字列、またはエラー文字列（"ERR:..."）
     */
    function callMainEngine(callExpr) {
        if (isBusy) { return "ERR:BUSY"; }
        isBusy = true;

        var holder = { value: null };
        try {
            var src = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { src += WORKER_FUNCS[i].toString(); }
            src += callExpr + ";";

            var bt = new BridgeTalk();
            bt.target = "illustrator";
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

    /**
     * 集計結果（OK|key=value|...）を解析する
     * @param {string} resp メインエンジンからの戻り値
     * @returns {object} キーと値のマップ（解析できない場合は null）
     */
    function parseCollect(resp) {
        if (!resp || resp.indexOf("OK|") !== 0) return null;
        var statPart = resp.substring(3);

        var map = {};
        var pairs = statPart.split("|");
        for (var i = 0; i < pairs.length; i++) {
            var eq = pairs[i].indexOf("=");
            if (eq > 0) { map[pairs[i].substring(0, eq)] = pairs[i].substring(eq + 1); }
        }
        if (map.docName != null) { try { map.docName = decodeURIComponent(map.docName); } catch (e) {} }

        return map;
    }

    /* ============================================================
       状態保持（常駐エンジン） / Session state (resident engine)
       ============================================================ */
    if (!$.global.__pathInspectorState) {
        $.global.__pathInspectorState = { location: null };
    }

    /* ============================================================
       パレット構築 / Build palette
       ============================================================ */

    /**
     * ラベルと値のペアを 1 行追加する
     * @param {object} panel 追加先のパネル
     * @param {string} labelText ラベル文字列
     * @param {number} labelWidth ラベルの幅（px）
     * @returns {object} 値表示用の statictext
     */
    function addStatRow(panel, labelText, labelWidth) {
        var row = panel.add("group");
        row.orientation = "row";
        var lbl = row.add("statictext", undefined, labelText);
        lbl.justify = "right";
        lbl.preferredSize.width = labelWidth;
        var val = row.add("statictext", undefined, "-");
        val.preferredSize.width = VALUE_WIDTH;
        return val;
    }

    /**
     * パレットを構築する
     * @returns {object} 構築済みの Window（palette）
     */
    function buildPalette() {
        var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
        win.orientation = "column";
        win.alignChildren = "center";
        win.margins = [15, 10, 15, 15];

        var content = win.add("group");
        content.orientation = "column";
        content.alignChildren = ["fill", "top"];
        content.margins = [10, 15, 10, 10];

        var v = {};

        var panelPath = content.add("panel", undefined, L('panel.path'));
        panelPath.orientation = "column";
        panelPath.alignChildren = ["fill", "top"];
        panelPath.margins = PANEL_MARGINS;

        v.pathCount = addStatRow(panelPath, L('row.pathCount'), LABEL_WIDTH);
        v.openPath = addStatRow(panelPath, L('row.openPath'), LABEL_WIDTH);
        v.closedPath = addStatRow(panelPath, L('row.closedPath'), LABEL_WIDTH);
        v.anchors = addStatRow(panelPath, L('row.anchors'), LABEL_WIDTH);
        v.handles = addStatRow(panelPath, L('row.handles'), LABEL_WIDTH);
        v.compoundPath = addStatRow(panelPath, L('row.compoundPath'), LABEL_WIDTH);
        v.compoundShape = addStatRow(panelPath, L('row.compoundShape'), LABEL_WIDTH);

        /* ステータス / Status line */
        var statusText = win.add("statictext", undefined, L('status.ready'));
        statusText.alignment = ["fill", "bottom"];

        /**
         * ステータス行を更新する
         * @param {string} msg 表示メッセージ
         * @returns {void}
         */
        function setStatus(msg) { try { statusText.text = msg; } catch (e) {} }

        /**
         * 集計値をパネルへ反映する
         * @param {object} m 集計結果のマップ
         * @returns {void}
         */
        function applyValues(m) {
            v.pathCount.text = m.pathCountSel + " / " + m.pathCountAll;
            v.openPath.text = m.openSel + " / " + m.openAll;
            v.closedPath.text = m.closedSel + " / " + m.closedAll;
            v.anchors.text = m.anchorSel + " / " + m.anchorAll;
            v.handles.text = m.handleSel + " / " + m.handleAll;
            v.compoundPath.text = m.cpathSel + " / " + m.cpathAll;
            v.compoundShape.text = m.cshapeSel + " / " + m.cshapeAll;
        }

        /**
         * 表示中の値をクリアする
         * @returns {void}
         */
        function clearValues() {
            for (var kk in v) { if (v.hasOwnProperty(kk)) { try { v[kk].text = "-"; } catch (e) {} } }
        }

        /**
         * メインエンジンへ集計を委譲し、結果を表示に反映する
         * @returns {void}
         */
        function refresh() {
            setStatus(L('status.busy'));
            var resp = callMainEngine("wkCollect()");

            if (resp === "ERR:BUSY") { setStatus(L('status.busy')); return; }
            if (resp === null || resp === "ERR:TIMEOUT") { setStatus(L('status.timeout')); return; }
            if (resp === "NODOC") { setStatus(L('status.noDoc')); clearValues(); return; }
            if (resp.indexOf("ERR:") === 0) { setStatus(L('status.error') + ": " + resp.substring(4)); return; }

            var map = parseCollect(resp);
            if (!map) { setStatus(L('status.error')); return; }

            applyValues(map);

            var selN = parseInt(map.selCount, 10) || 0;
            if (selN > 0) {
                setStatus(L('status.selectedPrefix') + selN + L('status.selectedSuffix'));
            } else {
                setStatus(L('status.wholeDoc'));
            }
        }

        /**
         * 集計結果をテキストファイルとしてデスクトップへ書き出す
         * @returns {void}
         */
        function exportReport() {
            setStatus(L('status.busy'));
            var resp = callMainEngine("wkCollect()");
            if (resp === "NODOC") { setStatus(L('status.noDoc')); return; }
            if (resp === null || resp === "ERR:TIMEOUT") { setStatus(L('status.timeout')); return; }
            if (typeof resp === "string" && resp.indexOf("ERR:") === 0) { setStatus(L('status.error') + ": " + resp.substring(4)); return; }

            var m = parseCollect(resp);
            if (!m) { setStatus(L('status.error')); return; }

            try {
                var fullName = m.docName || "";
                var baseName = fullName.replace(/\.[^\.]+$/, "");
                var today = new Date();
                var yyyy = today.getFullYear();
                var mm = ("0" + (today.getMonth() + 1)).slice(-2);
                var dd = ("0" + today.getDate()).slice(-2);
                var dateStr = yyyy + mm + dd;

                var path = Folder.desktop + "/path-" + baseName + "-" + dateStr + ".txt";
                var file = new File(path);

                /**
                 * 「選択 / 全体」形式で 1 行書き出す
                 * @param {string} path2 ラベルのドットパス
                 * @param {string} selVal 選択側の値
                 * @param {string} allVal 全体側の値
                 * @returns {void}
                 */
                function wPair(path2, selVal, allVal) { file.writeln(LX(path2) + " " + selVal + " / " + allVal); }

                /**
                 * セクション見出しを書き出す
                 * @param {string} path2 見出しのドットパス
                 * @returns {void}
                 */
                function wSection(path2) { file.writeln(""); file.writeln(L(path2)); }

                if (file.open("w")) {
                    file.writeln(L('report.title'));
                    file.writeln(L('report.document') + " " + fullName);
                    file.writeln(L('report.date') + " " + yyyy + "-" + mm + "-" + dd);
                    file.writeln("");
                    file.writeln(L('report.valueNote'));

                    wSection('section.paths');
                    wPair('row.pathCount', m.pathCountSel, m.pathCountAll);
                    wPair('row.openPath', m.openSel, m.openAll);
                    wPair('row.closedPath', m.closedSel, m.closedAll);
                    wPair('row.anchors', m.anchorSel, m.anchorAll);
                    wPair('row.handles', m.handleSel, m.handleAll);
                    wPair('row.compoundPath', m.cpathSel, m.cpathAll);
                    wPair('row.compoundShape', m.cshapeSel, m.cshapeAll);

                    file.close();
                    setStatus(L('status.exportedPrefix') + path);
                } else {
                    setStatus(L('status.exportFailOpen'));
                }
            } catch (err) {
                setStatus(L('status.error') + ": " + err);
            }
        }

        /* --- ボタン行 / Button row --- */
        var btnRow = win.add("group");
        btnRow.orientation = "row";
        btnRow.alignChildren = ["fill", "center"];
        btnRow.alignment = ["fill", "bottom"];

        var btnLeft = btnRow.add("group");
        btnLeft.alignChildren = ["left", "center"];
        var btnExport = btnLeft.add("button", undefined, L('button.exportPreset'));
        btnExport.helpTip = L('hint.esc');

        var spacer = btnRow.add("statictext", undefined, "");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRight = btnRow.add("group");
        btnRight.alignChildren = ["right", "center"];
        var btnRefresh = btnRight.add("button", undefined, L('button.refresh'));
        btnRefresh.helpTip = L('hint.refresh') + "\n" + L('hint.esc');

        btnExport.onClick = exportReport;
        btnRefresh.onClick = refresh;

        /* キー操作 / Key handling
           Esc: 閉じる / close
           ⌘R: 更新 / Cmd+R refresh */
        win.addEventListener("keydown", function (ev) {
            var key = "";
            try { key = ev && ev.keyName ? String(ev.keyName).toUpperCase() : ""; } catch (e) { key = ""; }

            if (key === "ESCAPE") {
                try { win.close(); } catch (e1) {}
            } else if (ev && ev.metaKey && key === "R") {
                refresh();
                try { if (ev.preventDefault) ev.preventDefault(); } catch (e2) {}
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
       位置の記憶・復元 / Remember & restore location
       ============================================================ */

    /**
     * 記憶した位置へパレットを復元する（未記憶なら中央）
     * @param {object} win 対象の Window
     * @returns {void}
     */
    function restoreLocation(win) {
        try {
            var loc = $.global.__pathInspectorState.location;
            if (loc && loc.length === 2) {
                win.location = [loc[0], loc[1]];
            } else {
                win.center();
            }
        } catch (e) {
            win.center();
        }
    }

    /**
     * パレットの現在位置を記憶する
     * @param {object} win 対象の Window
     * @returns {void}
     */
    function rememberLocation(win) {
        try {
            if (win.location && win.location.length === 2) {
                $.global.__pathInspectorState.location = [win.location[0], win.location[1]];
            }
        } catch (e) {}
    }

    /* ============================================================
       起動 / Entry point
       ============================================================ */

    /**
     * パレットを表示する（多重起動時は既存を閉じてから再表示）
     * @returns {void}
     */
    function showPalette() {
        if ($.global.__PathInspectorPalette) {
            try { $.global.__PathInspectorPalette.close(); } catch (e) {}
            $.global.__PathInspectorPalette = null;
        }

        var win = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__PathInspectorPalette = win;
        win.onClose = function () {
            rememberLocation(win);
            try { app.redraw(); } catch (e) {}
            try { $.global.__PathInspectorPalette = null; } catch (e2) {}
        };

        restoreLocation(win);
        win.show();
    }

    showPalette();

})();
