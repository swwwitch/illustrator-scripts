#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* Illustrator ExtendScript
   選択テキストの整形（ダイアログ）
   更新日：2026-02-23
   - 「削除」パネル、「ナンバリング」パネル、「改行」パネル、「アルファベット」パネル、「その他」パネルを使用
   - 各 panel の margins をすべて [15, 20, 15, 10] に設定
   - 機能：行頭/行末スペース削除、ナンバリング削除、強制改行↔改行 変換、空行の整理（連続改行の圧縮）、タブ→半角スペース、ナンバリングの振り直し、全角英数字を半角に
   - ナンバリング：［リセット］で（プレビューON時）ダイアログ初期状態へ復元して再プレビュー
*/

(function () {
    if (app.documents.length === 0) { alert("ドキュメントが開かれていません。"); return; }
    var sel = app.selection;

    // 選択がない場合は、ドキュメント内すべてのテキストを対象にする
    if (!sel || sel.length === 0) {
        sel = [];
        var allFrames = app.activeDocument.textFrames;
        for (var i = 0; i < allFrames.length; i++) {
            sel.push(allFrames[i]);
        }
    }

    function isTextFrame(o) { return o && o.typename === "TextFrame"; }
    function isTextRange(o) { return o && o.typename === "TextRange"; }


    // 英字のケース変換モード（null=未選択）
    var __caseMode = null; // "upper" | "lower" | "word" | "sentence" | "title"

    // 例表示用のベーステキスト（ダイアログ表示時点の内容）
    var __exampleBaseText = "";

    function __normalizeExampleText(s) {
        if (s == null) return "";
        // 改行/強制改行をスペースに
        s = String(s).replace(/[\r\n\u0003]+/g, " ");
        // 連続スペースを軽く整理
        s = s.replace(/[ \u3000\t]+/g, " ").replace(/^\s+|\s+$/g, "");
        // 長すぎるとUIが伸びるので短縮
        if (s.length > 40) s = s.substring(0, 40) + "…";
        return s;
    }

    function __getFirstSelectedTextSnapshot() {
        // ベースラインが取れていれば、その最初の TextRange の baseline を使う
        if (__baselineReady && __baselineRanges && __baselineRanges.length > 0) {
            try { return __baselineRanges[0].baseline; } catch (_) { }
        }
        // フォールバック：選択（または全テキスト）から最初の TextRange の現在値
        try {
            var tmp = [];
            for (var i = 0; i < sel.length; i++) collectTextRanges(sel[i], tmp);
            if (tmp.length > 0) return tmp[0].contents;
        } catch (_) { }
        return "";
    }

    function __formatCaseExample(mode, s) {
        try {
            if (mode === "upper") return s.toUpperCase();
            if (mode === "lower") return s.toLowerCase();
            if (mode === "word") return toWordCap(s);
            if (mode === "sentence") return toSentenceCasePreserveAcronyms(s);
            if (mode === "title") return toTitleCase(s);
        } catch (_) { }
        return s;
    }

    // 半角/全角数字（ナンバリング用）
    var __digitsForRenumber = "0-9\uFF10-\uFF19";

    // ナンバリング振り直しの出力スタイル
    // "num"=1. , "alpha"=A. , "circled"=① , "dot"=・ , "hyphen"=-
    var __renumberStyle = "num";

    // ダイアログ表示時点（初回プレビュー適用後）の状態に戻すためのベースライン
    var __baselineRanges = []; // { tr: TextRange, baseline: String }
    var __baselineReady = false;

    function __addBaseline(tr) {
        for (var i = 0; i < __baselineRanges.length; i++) {
            if (__baselineRanges[i].tr === tr) return;
        }
        try { __baselineRanges.push({ tr: tr, baseline: tr.contents }); } catch (_) { }
    }

    function __collectBaseline(item) {
        if (!item) return;
        if (isTextRange(item)) { __addBaseline(item); return; }
        if (isTextFrame(item)) { __addBaseline(item.textRange); return; }
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) __collectBaseline(item.pageItems[i]);
        }
    }

    function __takeBaseline() {
        __baselineRanges = [];
        for (var i = 0; i < sel.length; i++) __collectBaseline(sel[i]);
        __baselineReady = true;
    }

    function __restoreBaseline() {
        if (!__baselineReady) return;
        for (var i = 0; i < __baselineRanges.length; i++) {
            try { __baselineRanges[i].tr.contents = __baselineRanges[i].baseline; } catch (_) { }
        }
    }


    // --- Dialog ---
    var w = new Window("dialog", "テキストのクリーンアップと変換");
    w.orientation = "column";
    w.alignChildren = ["fill", "top"];
    w.preferredSize.width = 500;

    // --- Tabs ---
    var tabs = w.add("tabbedpanel");
    tabs.alignChildren = ["fill", "top"];
    tabs.preferredSize.width = 500;

    // margins 配列（left, top, right, bottom）
    var panelMargins = [15, 20, 15, 10];

    // --- 削除（Tab） ---
    var delPanel = tabs.add("tab", undefined, "削除");
    delPanel.orientation = "column";
    delPanel.alignChildren = ["left", "top"];
    try { delPanel.margins = panelMargins; } catch (_) { }

    // 行頭/行末
    var cbLead = delPanel.add("checkbox", undefined, "行頭のスペース");
    cbLead.value = true;
    var cbTrail = delPanel.add("checkbox", undefined, "行末のスペース");
    cbTrail.value = true;

    // 連続スペース
    var cbMulti = delPanel.add("checkbox", undefined, "連続するスペース");
    cbMulti.value = true;

    // --- ナンバリング（Tab） ---
    var numPanel = tabs.add("tab", undefined, "ナンバリング");
    numPanel.orientation = "column";
    numPanel.alignChildren = ["left", "top"];
    try { numPanel.margins = panelMargins; } catch (_) { }

    // 2カラム（左：ボタン / 右：形式）
    var numCols = numPanel.add("group");
    numCols.orientation = "row";
    numCols.alignChildren = ["fill", "top"];

    var numColLeft = numCols.add("group");
    numColLeft.orientation = "column";
    numColLeft.alignChildren = ["left", "top"];

    var numColRight = numCols.add("group");
    numColRight.orientation = "column";
    numColRight.alignChildren = ["left", "top"];

    // ナンバリング（状態フラグ）
    var cbNumEnd = { value: false };     // ナンバリング削除
    var cbRenumber = { value: false };   // ナンバリングの振り直し

    var numBtnGroup = numColLeft.add("group");
    numBtnGroup.orientation = "column";
    numBtnGroup.alignChildren = ["left", "top"];

    function makeSmallActionButton(parent, label, onRun) {
        var b = parent.add("button", undefined, label);
        b.preferredSize.height = 22; // 少し小さめ
        b.onClick = function () {
            try { onRun(); } catch (_) { }
            requestPreview();
        };
        return b;
    }

    makeSmallActionButton(numBtnGroup, "行頭マーカー削除", function () {
        cbNumEnd.value = true;
        cbRenumber.value = false;
    });

    makeSmallActionButton(numBtnGroup, "ナンバリングの振り直し", function () {
        cbRenumber.value = true;
        cbNumEnd.value = false;
    });

    var btnResetNumbering = numBtnGroup.add("button", undefined, "リセット");
    btnResetNumbering.preferredSize.height = 22;
    btnResetNumbering.onClick = function () {
        cbNumEnd.value = false;
        cbRenumber.value = false;
        __renumberStyle = "num";
        try { rbNumDot.value = true; } catch (_) { }

        // プレビューON時：ナンバリング処理で失われた元の行頭マーカー等を戻すため、
        // ダイアログ表示時点（初回プレビュー適用後）のベースラインへ復元してから再適用する。
        if (chkPreview && chkPreview.value) {
            __restoreBaseline();
        }

        requestPreview();
        try { updateCaseExamples(); } catch (_) { }
    };

    // 出力形式（振り直し）
    var numStylePanel = numColRight.add("panel", undefined, "形式");
    numStylePanel.orientation = "column";
    numStylePanel.alignChildren = ["left", "top"];
    try { numStylePanel.margins = panelMargins; } catch (_) { }

    var rbNumDot = numStylePanel.add("radiobutton", undefined, "1. いちご");
    var rbAlphaDot = numStylePanel.add("radiobutton", undefined, "A. いちご");
    var rbCircled = numStylePanel.add("radiobutton", undefined, "① いちご");
    var rbDot = numStylePanel.add("radiobutton", undefined, "・いちご");
    var rbHyphen = numStylePanel.add("radiobutton", undefined, "- いちご");

    // デフォルト: 1. 形式
    rbNumDot.value = true;
    __renumberStyle = "num";

    function __setRenumberStyle(style) {
        __renumberStyle = style;
        // requestPreview();
    }

    rbNumDot.onClick = function () { __setRenumberStyle("num"); };
    rbAlphaDot.onClick = function () { __setRenumberStyle("alpha"); };
    rbCircled.onClick = function () { __setRenumberStyle("circled"); };
    rbDot.onClick = function () { __setRenumberStyle("dot"); };
    rbHyphen.onClick = function () { __setRenumberStyle("hyphen"); };

    // --- 行（Tab） ---
    var optPanel = tabs.add("tab", undefined, "行");
    optPanel.orientation = "column";
    optPanel.alignChildren = ["left", "top"];
    try { optPanel.margins = panelMargins; } catch (_) { }

    // 改行変換（状態フラグ）
    var rbForcedToPara = { value: false };   // 強制改行→改行
    var rbParaToForced = { value: false };   // 改行→強制改行

    var brBtnGroup = optPanel.add("group");
    brBtnGroup.orientation = "column";
    brBtnGroup.alignChildren = ["left", "top"];

    function makeSmallBrButton(parent, label, onRun) {
        var b = parent.add("button", undefined, label);
        b.preferredSize.height = 22; // 少し小さめ
        b.onClick = function () {
            try { onRun(); } catch (_) { }
            requestPreview();
        };
        return b;
    }

    makeSmallBrButton(brBtnGroup, "強制改行を改行に", function () {
        rbForcedToPara.value = true;
        rbParaToForced.value = false;
    });

    makeSmallBrButton(brBtnGroup, "改行を強制改行に", function () {
        rbParaToForced.value = true;
        rbForcedToPara.value = false;
    });

    // 空行整理（状態フラグ）
    var cbCompressBlank = { value: false };

    var btnCompressBlank = optPanel.add("button", undefined, "空行の整理（連続改行の圧縮）");
    btnCompressBlank.preferredSize.height = 22; // 少し小さめ
    btnCompressBlank.onClick = function () {
        cbCompressBlank.value = true;
        requestPreview();
    };
    
    // --- ソート（Panel） ---
    var sortPanel = optPanel.add("panel", undefined, "ソート");
    sortPanel.orientation = "column";
    sortPanel.alignChildren = ["left", "top"];
    try { sortPanel.margins = panelMargins; } catch (_) { }

    var sortGroup = sortPanel.add("group");
    sortGroup.orientation = "column";
    sortGroup.alignChildren = ["left", "top"];

    var rbSortAsc = sortGroup.add("radiobutton", undefined, "ソート");
    var rbReverse = sortGroup.add("radiobutton", undefined, "行を逆順に");
    var rbUniqueAdjacent = sortGroup.add("radiobutton", undefined, "隣接する重複行を削除");

    rbSortAsc.value = false;
    rbReverse.value = false;
    rbUniqueAdjacent.value = false;

    // --- アルファベット（Tab） ---
    var alnumPanel = tabs.add("tab", undefined, "アルファベット");
    alnumPanel.orientation = "column";
    alnumPanel.alignChildren = ["left", "top"];
    try { alnumPanel.margins = panelMargins; } catch (_) { }


    // --- 英字のケース変換（ボタン＋例） ---
    var caseGroup = alnumPanel.add("group");
    caseGroup.orientation = "column";
    caseGroup.alignChildren = ["left", "top"];

    var __caseExampleLabels = {}; // mode -> statictext

    function makeSmallButton(parent, label, mode) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];

        var b = row.add("button", undefined, label);
        b.preferredSize.height = 22; // 少し小さめ
        b.preferredSize.width = 220;
        b.onClick = function () {
            __caseMode = mode;
            requestPreview();
        };

        var st = row.add("statictext", undefined, "");
        st.preferredSize.width = 240;
        try { st.justify = "left"; } catch (_) { }

        __caseExampleLabels[mode] = st;
        return b;
    }

    function updateCaseExamples() {
        __exampleBaseText = __normalizeExampleText(__getFirstSelectedTextSnapshot());
        var src = __exampleBaseText;
        try {
            if (__caseExampleLabels.upper) __caseExampleLabels.upper.text = __normalizeExampleText(__formatCaseExample("upper", src));
            if (__caseExampleLabels.lower) __caseExampleLabels.lower.text = __normalizeExampleText(__formatCaseExample("lower", src));
            if (__caseExampleLabels.word) __caseExampleLabels.word.text = __normalizeExampleText(__formatCaseExample("word", src));
            if (__caseExampleLabels.sentence) __caseExampleLabels.sentence.text = __normalizeExampleText(__formatCaseExample("sentence", src));
            if (__caseExampleLabels.title) __caseExampleLabels.title.text = __normalizeExampleText(__formatCaseExample("title", src));
        } catch (_) { }
    }

    makeSmallButton(caseGroup, "すべて大文字に", "upper");
    makeSmallButton(caseGroup, "すべて小文字に", "lower");
    makeSmallButton(caseGroup, "単語の先頭のみ大文字", "word");
    makeSmallButton(caseGroup, "文頭のみ大文字", "sentence");
    makeSmallButton(caseGroup, "英語タイトル形式", "title");

    var btnResetCase = caseGroup.add("button", undefined, "リセット");
    btnResetCase.preferredSize.height = 22; // 少し小さめ
    btnResetCase.onClick = function () {
        // テキストを「ダイアログを開いた時点」に戻す
        __restoreBaseline();

        // UI を初期値へ
        try {
            cbLead.value = true;
            cbTrail.value = true;
            cbMulti.value = true;

            cbNumEnd.value = false;
            cbRenumber.value = false;
            __renumberStyle = "num";
            try { rbNumDot.value = true; } catch (_) { }

            rbForcedToPara.value = false;
            rbParaToForced.value = false;
            cbCompressBlank.value = false;

            cbTabToSpace.value = false;
            cbZenkakuAlnumToHankaku.value = false;
            cbHyphenToUnderscore.value = false;
            cbUnderscoreToHyphen.value = false;

            rbSortAsc.value = false;
            rbReverse.value = false;
            rbUniqueAdjacent.value = false;

            __caseMode = null;

            chkPreview.value = true;
        } catch (_) { }

        // 初期状態（=プレビューON前提）として再描画のみ
        try { app.redraw(); } catch (_) { }
        try { updateCaseExamples(); } catch (_) { }
    };

    // --- その他（Tab） ---
    var otherPanel = tabs.add("tab", undefined, "その他");
    otherPanel.orientation = "column";
    otherPanel.alignChildren = ["left", "top"];
    try { otherPanel.margins = panelMargins; } catch (_) { }

    var cbTabToSpace = otherPanel.add("checkbox", undefined, "タブ→半角スペースに");
    cbTabToSpace.value = false;

    var cbZenkakuAlnumToHankaku = otherPanel.add("checkbox", undefined, "全角英数字を半角に");
    cbZenkakuAlnumToHankaku.value = false;

    var cbHyphenToUnderscore = otherPanel.add("checkbox", undefined, "ハイフンをアンダースコアに");
    cbHyphenToUnderscore.value = false;

    var cbUnderscoreToHyphen = otherPanel.add("checkbox", undefined, "アンダースコアをハイフンに");
    cbUnderscoreToHyphen.value = false;

    rbSortAsc.value = false;
    rbReverse.value = false;
    rbUniqueAdjacent.value = false;

    // Footer (Preview + Buttons)
    var footer = w.add("group");
    footer.orientation = "row";
    footer.alignChildren = ["fill", "center"];

    var gLeft = footer.add("group");
    gLeft.orientation = "row";
    gLeft.alignChildren = ["left", "center"];
    gLeft.alignment = ["left", "center"];

    var chkPreview = gLeft.add("checkbox", undefined, "プレビュー");
    chkPreview.value = true;

    var gRight = footer.add("group");
    gRight.orientation = "row";
    gRight.alignChildren = ["right", "center"];
    gRight.alignment = ["right", "center"];

    var cancel = gRight.add("button", undefined, "キャンセル", { name: "cancel" });
    var ok = gRight.add("button", undefined, "OK", { name: "ok" });

    function isAnyOptionSelected() {
        return (
            cbLead.value || cbTrail.value || cbNumEnd.value ||
            rbForcedToPara.value || rbParaToForced.value ||
            cbCompressBlank.value || cbTabToSpace.value || cbRenumber.value ||
            cbZenkakuAlnumToHankaku.value ||
            (__caseMode != null) ||
            cbHyphenToUnderscore.value || cbUnderscoreToHyphen.value ||
            rbSortAsc.value || rbReverse.value || rbUniqueAdjacent.value
        );
    }

    ok.onClick = function () {
        if (!isAnyOptionSelected()) {
            alert("実行する処理を1つ以上チェックしてください。");
            return;
        }
        // 簡易プレビュー方式：
        // - プレビューONなら既に適用済みなので、ここでは何もしない
        // - プレビューOFFなら、OK時に1回だけ適用
        if (!chkPreview.value) {
            try { applyProcessToSelection(); } catch (_) { }
        }
        w.close(1);
    };
    cancel.onClick = function () {
        w.close(0);
    };

    // --- Preview wiring (simple / history will be polluted) ---
    function __pushUniqueTextRange(list, tr) {
        for (var i = 0; i < list.length; i++) {
            if (list[i] === tr) return;
        }
        list.push(tr);
    }

    function collectTextRanges(item, list) {
        if (!item) return;
        if (isTextRange(item)) { __pushUniqueTextRange(list, item); return; }
        if (isTextFrame(item)) { __pushUniqueTextRange(list, item.textRange); return; }
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) collectTextRanges(item.pageItems[i], list);
        }
    }

    function applyProcessToSelection() {
        var targets = [];
        for (var i = 0; i < sel.length; i++) collectTextRanges(sel[i], targets);
        for (var j = 0; j < targets.length; j++) formatTextRange(targets[j]);
    }

    function requestPreview() {
        if (!chkPreview.value) return;
        try {
            applyProcessToSelection();
            app.redraw(); // プレビューON時のみ強制再描画
        } catch (_) { }
    }

    // 初回表示時：プレビューがONならまず適用 → その結果を「ベースライン」として記録
    w.onShow = function () {
        // 初回表示時：プレビューがONならまず適用 → その結果を「ベースライン」として記録
        if (chkPreview.value) {
            requestPreview();
        }
        // requestPreview 内で apply が走った後の状態（または未適用のまま）を保持
        __takeBaseline();
        try { updateCaseExamples(); } catch (_) { }
    };

    // UI 変更時にプレビュー更新
    cbLead.onClick = requestPreview;
    cbTrail.onClick = requestPreview;
    cbMulti.onClick = requestPreview;

    // (cbNumEnd/cbRenumber are now state objects, not checkboxes)

    // (rbForcedToPara, rbParaToForced now use button logic; no onClick here)

    cbTabToSpace.onClick = requestPreview;
    cbZenkakuAlnumToHankaku.onClick = requestPreview;
    cbHyphenToUnderscore.onClick = requestPreview;
    cbUnderscoreToHyphen.onClick = requestPreview;


    rbSortAsc.onClick = requestPreview;
    rbReverse.onClick = requestPreview;
    rbUniqueAdjacent.onClick = requestPreview;

    chkPreview.onClick = function () {
        // OFF にしても復元はしない（簡易プレビュー）
        if (chkPreview.value) requestPreview();
    };

    if (w.show() !== 1) return;

    // --- regex for spaces ---
    // 半角スペース + 全角スペース（常に対象）
    var spaces = " \u3000";
    var quant = cbMulti.value ? "+" : "";
    var reLead = new RegExp("(^|[\\r\\n\\u0003])[" + spaces + "]" + quant, "g");
    var reTrail = new RegExp("[" + spaces + "]" + quant + "([\\r\\n\\u0003]|$)", "g");

    // --- regex for line-head numbering ---
    // 行頭のみを対象にする（例："1" / "2." / "2. " など）
    var digits = "0-9\uFF10-\uFF19"; // 半角/全角数字
    // 行頭（または改行直後） + 数字 + （任意で . ） + （任意でスペース）
    var reHeadNumbering = new RegExp("(^|[\\r\\n\\u0003])[" + digits + "]+(?:[\\.\\uFF0E])?[" + spaces + "]*", "g");

    // --- 改行変換の定義 ---
    // Illustrator の TextRange.contents では、段落改行は "\r"。
    // 強制改行（ソフトリターン / Shift+Return）は "\u0003"（ETX）として入ることが多い。
    // 環境/経路によっては "\n" が混在する場合もあるため、両方を強制改行として扱う。
    var FORCED_BR = "\u0003"; // 強制改行（ETX）
    var reForced = /[\n\u0003]/g; // 強制改行（ETX）および混在するLF
    var rePara = /\r/g;           // 段落改行（CR）

    // 連続する空行は「残さない」：2つ以上の段落改行（空白だけの行を含む）を 1つの段落改行（\r）に圧縮
    var reBlankPara2Plus = /\r(?:[ \u3000\t]*\r)+/g;

    // --- タブ変換 ---
    var reTab = /\t/g;

    // --- ハイフン／アンダースコア変換 ---
    var reHyphen = /-/g;
    var reUnderscore = /_/g;

    // --- ナンバリングの振り直し ---
    // 行頭の「数字 + 任意の . + 任意のスペース」を連番に置換する。
    // 例: "1" / "2." / "2. " を 1,2,3... に振り直す（. の有無は元を踏襲）。

    function renumberLineHeadNumbering(s) {
        // 改行区切り（CR/ETX/LF）を保持して処理する
        var parts = s.split(/(\r|\u0003|\n)/);
        var n = 1;

        for (var i = 0; i < parts.length; i += 2) {
            var line = parts[i];
            if (line == null) continue;

            // 行頭インデント（半角/全角スペース）
            var indentMatch = line.match(/^[ \u3000]*/);
            var indent = indentMatch ? indentMatch[0] : "";
            var rest = line.substring(indent.length);

            // 既存の先頭マーカーを除去（番号/ドット/中黒/ハイフン等）
            rest = rest.replace(new RegExp(
                "^(?:" +
                "[" + __digitsForRenumber + "]+(?:[\\.\\uFF0E])?" +
                "|[A-Za-z](?:[\\.\\uFF0E])" +
                "|[\\u2460-\\u2473]" +
                "|・" +
                "|-" +
                ")[ \\u3000]*"), "");

            // 残りが空（空白のみ含む）なら、番号は振らずにそのまま
            if (/^[ \u3000\t]*$/.test(rest)) {
                parts[i] = indent + rest;
                continue;
            }

            // 先頭の余計な空白は落として、出力形式を選択
            rest = rest.replace(/^[ \u3000\t]+/, "");
            // 出力形式を選択
            var prefix = "";
            if (__renumberStyle === "alpha") {
                // A, B, ... Z, AA, AB ...
                var k = n;
                var letters = "";
                while (k > 0) {
                    k--; // 1-based
                    letters = String.fromCharCode(65 + (k % 26)) + letters;
                    k = Math.floor(k / 26);
                }
                prefix = letters + ". ";
            } else if (__renumberStyle === "circled") {
                // ①(1)〜⑳(20) まで対応。超えたら (n)
                if (n >= 1 && n <= 20) {
                    prefix = String.fromCharCode(0x2460 + (n - 1)) + " ";
                } else {
                    prefix = "(" + String(n) + ") ";
                }
            } else if (__renumberStyle === "dot") {
                prefix = "・";
            } else if (__renumberStyle === "hyphen") {
                prefix = "- ";
            } else {
                // default: 1.
                prefix = String(n) + ". ";
            }

            parts[i] = indent + prefix + rest;
            n++;
        }

        return parts.join("");
    }


    // --- 全角英数字 → 半角 ---
    // 対象: 全角 0-9（FF10-FF19）, A-Z（FF21-FF3A）, a-z（FF41-FF5A）
    function zenkakuAlnumToHankaku(s) {
        var out = "";
        for (var i = 0; i < s.length; i++) {
            var c = s.charCodeAt(i);
            var isZNum = (c >= 0xFF10 && c <= 0xFF19);
            var isZAZ = (c >= 0xFF21 && c <= 0xFF3A);
            var isZaz = (c >= 0xFF41 && c <= 0xFF5A);
            if (isZNum || isZAZ || isZaz) {
                out += String.fromCharCode(c - 0xFEE0);
            } else {
                out += s.charAt(i);
            }
        }
        return out;
    }

    function toWordCap(s) {
        return s.replace(/\b([a-z])/g, function (_, c) { return c.toUpperCase(); });
    }

    function toSentenceCasePreserveAcronyms(s) {
        // 1) まず全文を小文字化するが、
        //    - NASA のような全大文字語
        //    - iPhone のような mixedCase / camelCase（途中に大文字を含む語）
        //    を退避して保護する
        var placeholders = [];
        function keep(m) {
            // toLowerCase() の影響を受けないよう、英字を含まないプレースホルダにする
            // Private Use Area をマーカーとして使う（通常の入力テキストと衝突しにくい）
            var key = "\uE000" + placeholders.length + "\uE001";
            placeholders.push(m);
            return key;
        }

        // 全大文字語（2文字以上）
        s = s.replace(/\b[A-Z]{2,}\b/g, keep);
        // mixedCase / camelCase（英字のみ・途中に大文字を含む）
        // 例: iPhone, eBay, PowerPoint, macOS, ChatGPT など
        s = s.replace(/\b[A-Za-z]*[a-z][A-Za-z]*[A-Z][A-Za-z]*\b/g, keep);

        s = s.toLowerCase();

        // 2) 文頭および .!? の後の英字を大文字に
        s = s.replace(/(^|[\.\!\?]\s+|[\r\n\u0003]+)([a-z])/g,
            function (_, prefix, c) { return prefix + c.toUpperCase(); });

        // 3) 退避した語を復元
        for (var i = 0; i < placeholders.length; i++) {
            var key = "\uE000" + i + "\uE001";
            // 置換対象が複数回出現しても確実に復元する
            while (s.indexOf(key) !== -1) {
                s = s.replace(key, placeholders[i]);
            }
        }

        return s;
    }

    // --- Title Case ---
    // John Gruber / John Resig の Title Caps を元にしたロジック（Illustrator向け移植版を参考）
    // 英字（A-Za-z）にのみ作用する。
    function toTitleCase(s) {
        // 小文字にする語（冠詞・接続詞・前置詞など）
        var small = "(a|abaft|aboard|about|above|absent|across|afore|after|against|along|alonside|amid|amidst|among|amongst|an|and|apopos|around|as|aside|astride|at|athwart|atop|barring|before|behind|below|beneath|beside|besides|between|betwixt|beyond|but|by|circa|concerning|despite|down|during|except|excluding|failing|following|for|from|given|in|including|inside|into|lest|like|mid|midst|minus|modula|near|next|nor|notwithstanding|of|off|on|onto|oppostie|or|out|outside|over|pace|per|plus|pro|qua|regarding|round|sans|save|than|that|the|through|throughout|till|times|to|toward|towards|under|underneath|unlike|until|unto|up|upon|versus|via|vice|with|within|without|worth|v[.]?|via|vs[.]?)";
        var punct = "([!\"#$%&'()*+,./:;<=>?@[\\\\\\]^_`{|}~-]*)";

        function lower(word) { return word.toLowerCase(); }
        function upper(word) { return word.substr(0, 1).toUpperCase() + word.substr(1); }

        function titleCaps(title) {
            var parts = [], split = /[:.;?!] |(?: |^)[\"\u00D2]/g, index = 0;

            while (true) {
                var m = split.exec(title);

                parts.push(
                    title.substring(index, m ? m.index : title.length)
                        .replace(/\b([A-Za-z][a-z.'\u00D5]*)\b/g, function (all) {
                            return /[A-Za-z]\.[A-Za-z]/.test(all) ? all : upper(all);
                        })
                        .replace(RegExp("\\b" + small + "\\b", "ig"), lower)
                        .replace(RegExp("^" + punct + small + "\\b", "ig"), function (all, p, word) {
                            return p + upper(word);
                        })
                        .replace(RegExp("\\b" + small + punct + "$", "ig"), upper)
                );

                index = split.lastIndex;

                if (m) parts.push(m[0]);
                else break;
            }

            return parts.join("")
                .replace(/ V(s?)\. /ig, " v$1. ")
                .replace(/(['\u00D5])S\b/ig, "$1s")
                .replace(/\b(AT&T|Q&A)\b/ig, function (all) { return all.toUpperCase(); });
        }

        return titleCaps(s);
    }

    // --- ソート／逆順／重複削除 ---
    function splitLinesCR(s) {
        return s.split("\r");
    }

    function joinLinesCR(lines) {
        return lines.join("\r");
    }

    function sortLinesAsc(s) {
        var lines = splitLinesCR(s);
        lines.sort();
        return joinLinesCR(lines);
    }

    function reverseLines(s) {
        var lines = splitLinesCR(s);
        lines.reverse();
        return joinLinesCR(lines);
    }

    function removeAdjacentDuplicates(s) {
        var lines = splitLinesCR(s);
        var result = [];
        for (var i = 0; i < lines.length; i++) {
            if (i === 0 || lines[i] !== lines[i - 1]) {
                result.push(lines[i]);
            }
        }
        return joinLinesCR(result);
    }

    // --- runtime regex builder ---
    // プレビュー中でも使えるように、必要な正規表現は都度（軽量に）組み立てる。
    function buildRuntimeRegexes() {
        // --- regex for spaces ---
        // 半角スペース + 全角スペース（常に対象）
        var spaces = " \u3000";
        var quant = cbMulti.value ? "+" : "";
        var reLead = new RegExp("(^|[\\r\\n\\u0003])[" + spaces + "]" + quant, "g");
        var reTrail = new RegExp("[" + spaces + "]" + quant + "([\\r\\n\\u0003]|$)", "g");

        // --- regex for line-head numbering ---
        var digits = "0-9\\uFF10-\\uFF19"; // 半角/全角数字
        var reHeadNumbering = new RegExp("(^|[\\r\\n\\u0003])[" + digits + "]+(?:[\\.\\uFF0E])?[" + spaces + "]*", "g");

        // --- regex for line-head markers (numbering/bullets) ---
        // 対象: 1. / 1 / A. / ① / ・ / - など（行頭 or 改行直後）
        var reHeadMarker = new RegExp(
            "(^|[\\r\\n\\u0003])" +
            "(?:" +
            "[" + digits + "]+(?:[\\.\\uFF0E])?" +           // 1 / 1. / １．
            "|[A-Za-z](?:[\\.\\uFF0E])" +                   // A. / b.
            "|[\\u2460-\\u2473]" +                           // ①-⑳
            "|・" +                                            // ・
            "|-" +                                             // -
            ")" +
            "[" + spaces + "\\t]*",                           // 直後の空白（半角/全角/タブ）
            "g"
        );

        // --- newline conversion ---
        var FORCED_BR = "\u0003";
        var reForced = /[\n\u0003]/g;
        var rePara = /\r/g;

        // --- blank line compression (remove blank lines completely) ---
        var reBlankPara2Plus = /\r(?:[ \u3000\t]*\r)+/g;

        // --- tabs ---
        var reTab = /\t/g;

        // --- hyphen / underscore ---
        var reHyphen = /-/g;
        var reUnderscore = /_/g;

        return {
            spaces: spaces,
            digits: digits,
            reLead: reLead,
            reTrail: reTrail,
            reHeadNumbering: reHeadNumbering,
            reHeadMarker: reHeadMarker,
            FORCED_BR: FORCED_BR,
            reForced: reForced,
            rePara: rePara,
            reBlankPara2Plus: reBlankPara2Plus,
            reTab: reTab,
            reHyphen: reHyphen,
            reUnderscore: reUnderscore
        };
    }

    function formatTextRange(tr) {
        var s = tr.contents;
        var rx = buildRuntimeRegexes();

        // 1) スペースの行頭・行末削除
        if (cbLead.value) {
            s = s.replace(rx.reLead, function (_m, p1) { return (p1 != null ? p1 : ""); });
        }
        if (cbTrail.value) {
            s = s.replace(rx.reTrail, function (_m, p1) { return (p1 != null ? p1 : ""); });
        }

        // 2) 行頭マーカー削除（番号/記号）
        if (cbNumEnd.value) {
            s = s.replace(rx.reHeadMarker, function (_m, p1) { return (p1 != null ? p1 : ""); });
        }

        // 3) 改行変換（順序注意）
        if (rbForcedToPara.value) {
            // 強制改行（ETX / LF）→ 段落改行（CR）
            s = s.replace(rx.reForced, "\r");
        } else if (rbParaToForced.value) {
            // 段落改行（CR）→ 強制改行（ETX）
            s = s.replace(rx.rePara, rx.FORCED_BR);
        }

        // 4) 空行の整理（連続改行の圧縮）
        if (cbCompressBlank.value) {
            // 連続する空行（空白だけの行を含む）を空行なしに圧縮（\r 1つに統合）
            s = s.replace(rx.reBlankPara2Plus, "\r");
        }

        // 5) タブ→半角スペース
        if (cbTabToSpace.value) {
            s = s.replace(rx.reTab, " ");
        }

        // 5.4) ハイフン／アンダースコア変換
        if (cbHyphenToUnderscore.value && !cbUnderscoreToHyphen.value) {
            s = s.replace(rx.reHyphen, "_");
        } else if (cbUnderscoreToHyphen.value && !cbHyphenToUnderscore.value) {
            s = s.replace(rx.reUnderscore, "-");
        }

        // 5.5) 全角英数字を半角に
        if (cbZenkakuAlnumToHankaku.value) {
            s = zenkakuAlnumToHankaku(s);
        }

        // 5.6) 英字のケース変換
        if (__caseMode === "upper") {
            s = s.toUpperCase();
        } else if (__caseMode === "lower") {
            s = s.toLowerCase();
        } else if (__caseMode === "word") {
            s = toWordCap(s);
        } else if (__caseMode === "sentence") {
            s = toSentenceCasePreserveAcronyms(s);
        } else if (__caseMode === "title") {
            s = toTitleCase(s);
        }

        // 5.7) ソート／逆順／重複削除
        if (rbSortAsc.value) {
            s = sortLinesAsc(s);
        } else if (rbReverse.value) {
            s = reverseLines(s);
        } else if (rbUniqueAdjacent.value) {
            s = removeAdjacentDuplicates(s);
        }

        // 6) ナンバリングの振り直し
        // ※ナンバリング削除と同時指定された場合は、削除を優先して振り直しは行わない
        if (cbRenumber.value && !cbNumEnd.value) {
            s = renumberLineHeadNumbering(s);
        }

        tr.contents = s;
    }



    // (post-dialog processing removed: handled by preview logic and ok.onClick)

    // alert("完了しました。");
})();
