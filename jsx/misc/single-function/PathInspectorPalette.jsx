#target illustrator
#targetengine "PathInspectorSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中またはドキュメント全体のパス統計を集計し、常駐パレットで表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspectorPalette.md

### Overview

Counts path statistics for the selection, or for the whole document, and shows them in a persistent palette.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspectorPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathInspectorPalette";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspectorPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspectorPalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

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

    /* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "パスのカウント", en: "Path Count" }
        },
        report: {
            title: { ja: "Path Inspector Report", en: "Path Inspector Report" },
            document: { ja: "ドキュメント", en: "Document" },
            date: { ja: "日付", en: "Date" },
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
            pathCount: { ja: "パス", en: "Paths" },
            openPath: { ja: "オープンパス", en: "Open Paths" },
            closedPath: { ja: "クローズパス", en: "Closed Paths" },
            anchors: { ja: "アンカーポイント", en: "Anchor Points" },
            handles: { ja: "ハンドル", en: "Handles" },
            compoundPath: { ja: "複合パス", en: "Compound Paths" },
            compoundShape: { ja: "複合シェイプ", en: "Compound Shapes" }
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
     * 書き出し用にラベル末尾のコロンを半角へ正規化する
     * @param {string} path ラベルのドットパス
     * @returns {string} 正規化済み文字列
     */
    function LX(path) {
        return getLabel(path) + ":";
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
            var currentSelection = doc.selection;
            if (!currentSelection) { currentSelection = []; }
            var selCount = currentSelection.length;

            var allItems = doc.pageItems;

            var cpathSel = 0, cpathAll = 0, cshapeSel = 0, cshapeAll = 0;

            for (var i = 0; i < currentSelection.length; i++) {
                if (currentSelection[i].typename === "CompoundPathItem") { cpathSel++; }
                if (currentSelection[i].typename === "PluginItem") {
                    try { if (currentSelection[i].name && currentSelection[i].name.indexOf("Compound Shape") !== -1) { cshapeSel++; } } catch (e) {}
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
            for (var i2 = 0; i2 < currentSelection.length; i2++) { wkCountPathStats(currentSelection[i2], pathStatsSel); }

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
            var win = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
            win.orientation = "column";
            win.alignChildren = "center";
            win.margins = [15, 10, 15, 15];

            var content = win.add("group");
            content.orientation = "column";
            content.alignChildren = ["fill", "top"];
            content.margins = [10, 15, 10, 10];

            var v = {};

            var panelPath = content.add("panel", undefined, getLabel('panel.path'));
            panelPath.orientation = "column";
            panelPath.alignChildren = ["fill", "top"];
            panelPath.margins = PANEL_MARGINS;

            v.pathCount = addStatRow(panelPath, labelText('row.pathCount'), LABEL_WIDTH);
            v.openPath = addStatRow(panelPath, labelText('row.openPath'), LABEL_WIDTH);
            v.closedPath = addStatRow(panelPath, labelText('row.closedPath'), LABEL_WIDTH);
            v.anchors = addStatRow(panelPath, labelText('row.anchors'), LABEL_WIDTH);
            v.handles = addStatRow(panelPath, labelText('row.handles'), LABEL_WIDTH);
            v.compoundPath = addStatRow(panelPath, labelText('row.compoundPath'), LABEL_WIDTH);
            v.compoundShape = addStatRow(panelPath, labelText('row.compoundShape'), LABEL_WIDTH);

            /* ステータス / Status line */
            var statusText = win.add("statictext", undefined, getLabel('status.ready'));
            statusText.alignment = ["fill", "bottom"];

            /**
             * ステータス行を更新する
             * @param {string} msg 表示メッセージ
             * @returns {void}
             */
            function setStatus(msg) { statusText.text = msg; }

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
                setStatus(getLabel('status.busy'));
                var resp = callMainEngine("wkCollect()");

                if (resp === "ERR:BUSY") { setStatus(getLabel('status.busy')); return; }
                if (resp === null || resp === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
                if (resp === "NODOC") { setStatus(getLabel('status.noDoc')); clearValues(); return; }
                if (resp.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', resp.substring(4))); return; }

                var map = parseCollect(resp);
                if (!map) { setStatus(getLabel('status.error')); return; }

                applyValues(map);

                var selN = parseInt(map.selCount, 10) || 0;
                if (selN > 0) {
                    setStatus(getLabel('status.selectedPrefix') + selN + getLabel('status.selectedSuffix'));
                } else {
                    setStatus(getLabel('status.wholeDoc'));
                }
            }

            /**
             * 集計結果をテキストファイルとしてデスクトップへ書き出す
             * @returns {void}
             */
            function exportReport() {
                setStatus(getLabel('status.busy'));
                var resp = callMainEngine("wkCollect()");
                if (resp === "NODOC") { setStatus(getLabel('status.noDoc')); return; }
                if (resp === null || resp === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
                if (typeof resp === "string" && resp.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', resp.substring(4))); return; }

                var m = parseCollect(resp);
                if (!m) { setStatus(getLabel('status.error')); return; }

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
                    function wSection(path2) { file.writeln(""); file.writeln(getLabel(path2)); }

                    if (file.open("w")) {
                        file.writeln(getLabel('report.title'));
                        file.writeln(labelText('report.document') + " " + fullName);
                        file.writeln(labelText('report.date') + " " + yyyy + "-" + mm + "-" + dd);
                        file.writeln("");
                        file.writeln(getLabel('report.valueNote'));

                        wSection('section.paths');
                        wPair('row.pathCount', m.pathCountSel, m.pathCountAll);
                        wPair('row.openPath', m.openSel, m.openAll);
                        wPair('row.closedPath', m.closedSel, m.closedAll);
                        wPair('row.anchors', m.anchorSel, m.anchorAll);
                        wPair('row.handles', m.handleSel, m.handleAll);
                        wPair('row.compoundPath', m.cpathSel, m.cpathAll);
                        wPair('row.compoundShape', m.cshapeSel, m.cshapeAll);

                        file.close();
                        setStatus(getLabel('status.exportedPrefix') + path);
                    } else {
                        setStatus(getLabel('status.exportFailOpen'));
                    }
                } catch (err) {
                    setStatus(labelValueText('status.error', err));
                }
            }

            /* --- ボタン行 / Button row --- */
            var btnRow = win.add("group");
            btnRow.orientation = "row";
            btnRow.alignChildren = ["fill", "center"];
            btnRow.alignment = ["fill", "bottom"];

            var btnLeft = btnRow.add("group");
            btnLeft.alignChildren = ["left", "center"];
            var btnExport = btnLeft.add("button", undefined, getLabel('button.exportPreset'));
            btnExport.helpTip = getLabel('hint.esc');

            var spacer = btnRow.add("statictext", undefined, "");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var btnRight = btnRow.add("group");
            btnRight.alignChildren = ["right", "center"];
            var btnRefresh = btnRight.add("button", undefined, getLabel('button.refresh'));
            btnRefresh.helpTip = getLabel('hint.refresh') + "\n" + getLabel('hint.esc');

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
                app.redraw();
                $.global.__PathInspectorPalette = null;
            };

            restoreLocation(win);
            win.show();
        }

        showPalette();

    })();

})();
