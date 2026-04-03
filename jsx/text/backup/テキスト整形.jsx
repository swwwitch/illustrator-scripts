#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* Illustrator ExtendScript
   選択テキストの整形（ダイアログ）
   更新日：2026-02-23
  - 「スペース」パネル、「ナンバリング」パネル、「改行」パネル、「アルファベット」パネル、「文字種」パネルを使用
   - 各 panel の margins をすべて [15, 20, 15, 10] に設定
   - UI：［スペース］タブ下部に［処理］［リセット］ボタンを配置（スペース削除（基本）/（応用）共通）
   - UI：［スペース］>［スペース削除（応用）］に「和文内の半角スペース」「和欧内の半角スペース」を追加（デフォルトOFF／和文内・和欧内とも実装）
   - UI：［スペース］>［スペース変換］に「アンダースコアに」「ハイフンに」を追加（デフォルトOFF）
   - 機能：行頭/行末スペース削除、連続スペース圧縮、和文内/和欧内の半角スペース削除、半角/全角スペース→アンダースコア/ハイフン変換、ナンバリング削除、強制改行↔改行 変換、空行の整理（連続改行の圧縮）、タブ→半角スペース、ナンバリングの振り直し、全角英数字を半角に
   - ナンバリング：［リセット］で（プレビューON時）ダイアログ初期状態へ復元して再プレビュー

   version 1.0.0（2026-02-23）
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
    // ［スペース削除］の［処理］ボタンで手動適用したか（OKボタンのバリデーション用）
    var __spaceManualApplied = false;

    // ［文字種］タブの［変換］ボタンで手動適用したか（OKボタンのバリデーション用）
    var __charTypeManualApplied = false;

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

    // --- スペース（Tab） ---
    var delPanel = tabs.add("tab", undefined, "スペース");
    delPanel.orientation = "column";
    delPanel.alignChildren = ["left", "top"];
    try { delPanel.margins = panelMargins; } catch (_) { }

    // スペース（Panel）
    var spacePanel = delPanel.add("panel", undefined, "スペース削除（基本）");
    spacePanel.orientation = "column";
    spacePanel.alignChildren = ["left", "top"];
    try { spacePanel.margins = panelMargins; } catch (_) { }

    // スペース（チェックボックス）
    var cbLead = spacePanel.add("checkbox", undefined, "行頭のスペース");
    cbLead.value = true;

    var cbTrail = spacePanel.add("checkbox", undefined, "行末のスペース");
    cbTrail.value = true;

    var cbMulti = spacePanel.add("checkbox", undefined, "連続するスペース");
    cbMulti.value = true;

    // スペース（応用Panel）
    var spaceAdvPanel = delPanel.add("panel", undefined, "スペース削除（応用）");
    spaceAdvPanel.orientation = "column";
    spaceAdvPanel.alignChildren = ["left", "top"];
    try { spaceAdvPanel.margins = panelMargins; } catch (_) { }

    var cbHalfSpaceWa = spaceAdvPanel.add("checkbox", undefined, "和文内の半角スペース");
    cbHalfSpaceWa.value = false;

    var cbHalfSpaceWaOu = spaceAdvPanel.add("checkbox", undefined, "和欧内の半角スペース");
    cbHalfSpaceWaOu.value = false;

    // スペース変換（Panel）
    var spaceConvPanel = delPanel.add("panel", undefined, "スペース変換");
    spaceConvPanel.orientation = "column";
    spaceConvPanel.alignChildren = ["left", "top"];
    try { spaceConvPanel.margins = panelMargins; } catch (_) { }

    // 変換（チェックボックス）
    // ※半角/全角スペースを _ または - に変換する
    var cbSpaceToUnderscore = spaceConvPanel.add("checkbox", undefined, "アンダースコアに");
    cbSpaceToUnderscore.value = false;

    var cbSpaceToHyphen = spaceConvPanel.add("checkbox", undefined, "ハイフンに");
    cbSpaceToHyphen.value = false;

    // スペース削除（ボタン）※基本/応用共通
    var gSpaceBtns = delPanel.add("group");
    gSpaceBtns.orientation = "row";
    gSpaceBtns.alignChildren = ["left", "center"];
    try { gSpaceBtns.margins = [0, 10, 0, 0]; } catch (_) { }

    var btnSpaceDo = gSpaceBtns.add("button", undefined, "処理");
    btnSpaceDo.preferredSize.height = 22;
    btnSpaceDo.onClick = function () {
        if (!(cbLead.value || cbTrail.value || cbMulti.value || cbHalfSpaceWa.value || cbHalfSpaceWaOu.value)) {
            alert("処理する項目を1つ以上選択してください。");
            return;
        }
        try {
            applySpaceOnlyToSelection();
            __spaceManualApplied = true;
            try { app.redraw(); } catch (_) { }
        } catch (_) { }
    };

    var btnSpaceReset = gSpaceBtns.add("button", undefined, "リセット");
    btnSpaceReset.preferredSize.height = 22;
    btnSpaceReset.onClick = function () {
        try {
            cbLead.value = true;
            cbTrail.value = true;
            cbMulti.value = true;
            cbHalfSpaceWa.value = false;
            cbHalfSpaceWaOu.value = false;
            __spaceManualApplied = false;
        } catch (_) { }
    };

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
        b.preferredSize.width = 150;
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
            cbHalfSpaceWa.value = false;
            cbHalfSpaceWaOu.value = false;

            cbNumEnd.value = false;
            cbRenumber.value = false;
            __renumberStyle = "num";
            try { rbNumDot.value = true; } catch (_) { }

            rbForcedToPara.value = false;
            rbParaToForced.value = false;
            cbCompressBlank.value = false;

            cbTabToSpace.value = false;
            cbZenkakuAlnumToHankaku.value = false;
            cbHiraToKata.value = false;
            cbKataToHira.value = false;
            cbKataToHalf.value = false;
            cbHalfToKata.value = false;
            cbSpaceToUnderscore.value = false;
            cbSpaceToHyphen.value = false;

            rbSortAsc.value = false;
            rbReverse.value = false;
            rbUniqueAdjacent.value = false;

            __caseMode = null;

            __charTypeManualApplied = false;
        } catch (_) { }

        // 初期状態（=プレビューON前提）として再描画のみ
        try { app.redraw(); } catch (_) { }
        try { updateCaseExamples(); } catch (_) { }
    };

    // --- 文字種（Tab） ---
    var otherPanel = tabs.add("tab", undefined, "文字種");
    otherPanel.orientation = "column";
    otherPanel.alignChildren = ["left", "top"];
    try { otherPanel.margins = panelMargins; } catch (_) { }

    var cbTabToSpace = otherPanel.add("checkbox", undefined, "タブ→半角スペースに");
    cbTabToSpace.value = false;



    var cbZenkakuAlnumToHankaku = otherPanel.add("checkbox", undefined, "全角英数字を半角に");
    cbZenkakuAlnumToHankaku.value = false;

    var cbHiraToKata = otherPanel.add("checkbox", undefined, "ひらがな→カタカナ");
    cbHiraToKata.value = false;

    var cbKataToHira = otherPanel.add("checkbox", undefined, "カタカナ→ひらがな");
    cbKataToHira.value = false;

    var cbKataToHalf = otherPanel.add("checkbox", undefined, "カタカナ→半角カタカナ");
    cbKataToHalf.value = false;

    var cbHalfToKata = otherPanel.add("checkbox", undefined, "半角カタカナ→全角カタカナ");
    cbHalfToKata.value = false;

    // 相互排他（チェックしても実行はしない）
    cbHiraToKata.onClick = function () { if (cbHiraToKata.value) cbKataToHira.value = false; };
    cbKataToHira.onClick = function () { if (cbKataToHira.value) cbHiraToKata.value = false; };
    cbKataToHalf.onClick = function () { if (cbKataToHalf.value) cbHalfToKata.value = false; };
    cbHalfToKata.onClick = function () { if (cbHalfToKata.value) cbKataToHalf.value = false; };

    // 文字種（ボタン）
    var gCharBtns = otherPanel.add("group");
    gCharBtns.orientation = "row";
    gCharBtns.alignChildren = ["left", "center"];
    try { gCharBtns.margins = [0, 10, 0, 0]; } catch (_) { }

    var btnCharDo = gCharBtns.add("button", undefined, "変換");
    btnCharDo.preferredSize.height = 22;
    btnCharDo.onClick = function () {
        if (!(cbTabToSpace.value || cbZenkakuAlnumToHankaku.value)) {
            alert("変換する項目を1つ以上選択してください。\n（タブ→半角スペース / 全角英数字を半角）");
            return;
        }
        try {
            applyCharTypeOnlyToSelection();
            __charTypeManualApplied = true;
            try { app.redraw(); } catch (_) { }
        } catch (_) { }
    };

    var btnCharReset = gCharBtns.add("button", undefined, "リセット");
    btnCharReset.preferredSize.height = 22;
    btnCharReset.onClick = function () {
        try {
            cbTabToSpace.value = false;
            cbZenkakuAlnumToHankaku.value = false;
            __charTypeManualApplied = false;
        } catch (_) { }
    };

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
            cbNumEnd.value ||
            rbForcedToPara.value || rbParaToForced.value ||
            cbCompressBlank.value || cbRenumber.value ||
            (__caseMode != null) ||
            cbSpaceToUnderscore.value || cbSpaceToHyphen.value ||
            rbSortAsc.value || rbReverse.value || rbUniqueAdjacent.value
        );
    }

    ok.onClick = function () {
        if (!isAnyOptionSelected() && !__spaceManualApplied && !__charTypeManualApplied) {
            alert("実行する処理を1つ以上チェックしてください。");
            return;
        }
        // 簡易プレビュー方式：
        // - プレビューONなら既に適用済みなので、ここでは何もしない
        // - プレビューOFFなら、OK時に1回だけ適用
        if (!chkPreview.value) {
            if (isAnyOptionSelected()) {
                try { applyProcessToSelection(); } catch (_) { }
            }
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

    // ［文字種］タブ用：文字種変換のみ適用
    function applyCharTypeOnlyToSelection() {
        var targets = [];
        for (var i = 0; i < sel.length; i++) collectTextRanges(sel[i], targets);
        for (var j = 0; j < targets.length; j++) formatCharTypeOnlyRange(targets[j]);
    }

    function formatCharTypeOnlyRange(tr) {
        var s = tr.contents;
        var rx = buildRuntimeRegexes();

        // タブ→半角スペース
        if (cbTabToSpace.value) {
            s = s.replace(rx.reTab, " ");
        }

        if (cbZenkakuAlnumToHankaku.value) {


            tr.contents = s;
        }

        // ［文字種］タブ用：文字種変換のみ適用
        function applyCharTypeOnlyToSelection() {
            var targets = [];
            for (var i = 0; i < sel.length; i++) collectTextRanges(sel[i], targets);
            for (var j = 0; j < targets.length; j++) formatCharTypeOnlyRange(targets[j]);
        }

        function formatCharTypeOnlyRange(tr) {
            var s = tr.contents;
            var rx = buildRuntimeRegexes();

            // タブ→半角スペース
            if (cbTabToSpace.value) {
                s = s.replace(rx.reTab, " ");
            }

            // 全角英数字→半角
            if (cbZenkakuAlnumToHankaku.value) {
                s = zenkakuAlnumToHankaku(s);
            }

            tr.contents = s;
        }

        // ［スペース］タブの「スペース削除」用：スペース関連のみ適用（基本/応用共通）
        function applySpaceOnlyToSelection() {
            var targets = [];
            for (var i = 0; i < sel.length; i++) collectTextRanges(sel[i], targets);
            for (var j = 0; j < targets.length; j++) formatSpaceOnlyRange(targets[j]);
        }

        function formatSpaceOnlyRange(tr) {
            var s = tr.contents;
            var rx = buildRuntimeRegexes();

            // 行頭/行末
            if (cbLead.value) {
                s = s.replace(rx.reLead, function (_m, p1) { return (p1 != null ? p1 : ""); });
            }
            if (cbTrail.value) {
                s = s.replace(rx.reTrail, function (_m, p1) { return (p1 != null ? p1 : ""); });
            }

            // 連続スペース
            if (cbMulti.value) {
                var parts = s.split(/(\r|\u0003|\n)/);
                for (var i = 0; i < parts.length; i += 2) {
                    var line = parts[i] != null ? String(parts[i]) : "";

                    // 行頭/行末のスペースは触らない（インデント等の保持）
                    var mLead = line.match(/^[ \u3000]+/);
                    var mTrail = line.match(/[ \u3000]+$/);
                    var lead = mLead ? mLead[0] : "";
                    var trail = mTrail ? mTrail[0] : "";

                    var start = lead.length;
                    var end = line.length - trail.length;
                    if (end < start) { start = 0; end = line.length; lead = ""; trail = ""; }

                    var core = line.substring(start, end);
                    core = core.replace(/[ \u3000]{2,}/g, " ");

                    parts[i] = lead + core + trail;
                }
                s = parts.join("");
            }

            // 和文内の半角スペース：英数字以外の文字の間にある半角スペースを削除
            // 例）"半角 スペース が" → "半角スペースが"
            // ※半角スペースのみ対象（全角スペースは対象外）
            // ※改行/強制改行をまたがないよう、行ごとに処理する
            if (cbHalfSpaceWa.value) {
                var ALNUM = "A-Za-z0-9\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5A"; // 半角/全角 英数字
                var reWaHalfSpace = new RegExp("([^" + ALNUM + "]) +(?=[^" + ALNUM + "])", "g");

                var parts2 = s.split(/(\r|\u0003|\n)/);
                for (var k = 0; k < parts2.length; k += 2) {
                    var line2 = parts2[k] != null ? String(parts2[k]) : "";

                    // 行頭/行末のスペースは触らない（インデント等の保持）
                    var mLead2 = line2.match(/^[ \u3000]+/);
                    var mTrail2 = line2.match(/[ \u3000]+$/);
                    var lead2 = mLead2 ? mLead2[0] : "";
                    var trail2 = mTrail2 ? mTrail2[0] : "";

                    var start2 = lead2.length;
                    var end2 = line2.length - trail2.length;
                    if (end2 < start2) { start2 = 0; end2 = line2.length; lead2 = ""; trail2 = ""; }

                    var core2 = line2.substring(start2, end2);
                    core2 = core2.replace(reWaHalfSpace, "$1");

                    parts2[k] = lead2 + core2 + trail2;
                }
                s = parts2.join("");
            }

            // 和欧内の半角スペース：和文（非ASCII）と欧文（ASCII）の間にある半角スペースを削除
            // 例）"日本語 ABC" → "日本語ABC" / "ABC 日本語" → "ABC日本語"
            // ※半角スペースのみ対象（全角スペースは対象外）
            // ※改行/強制改行をまたがないよう、行ごとに処理する
            // 正規表現の意図：
            //   ([^[ -~]]) ([ -~])  →  \1\2
            //   ([ -~]) ([^[ -~]])  →  \1\2
            if (cbHalfSpaceWaOu.value) {
                var reWaOuLeft = /([^ -~]) +([ -~])/g;
                var reWaOuRight = /([ -~]) +([^ -~])/g;

                var parts3 = s.split(/(\r|\u0003|\n)/);
                for (var t = 0; t < parts3.length; t += 2) {
                    var line3 = parts3[t] != null ? String(parts3[t]) : "";

                    // 行頭/行末のスペースは触らない（インデント等の保持）
                    var mLead3 = line3.match(/^[ \u3000]+/);
                    var mTrail3 = line3.match(/[ \u3000]+$/);
                    var lead3 = mLead3 ? mLead3[0] : "";
                    var trail3 = mTrail3 ? mTrail3[0] : "";

                    var start3 = lead3.length;
                    var end3 = line3.length - trail3.length;
                    if (end3 < start3) { start3 = 0; end3 = line3.length; lead3 = ""; trail3 = ""; }

                    var core3 = line3.substring(start3, end3);
                    core3 = core3.replace(reWaOuLeft, "$1$2");
                    core3 = core3.replace(reWaOuRight, "$1$2");

                    parts3[t] = lead3 + core3 + trail3;
                }
                s = parts3.join("");
            }

            tr.contents = s;
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

        // (cbNumEnd/cbRenumber are now state objects, not checkboxes)

        // (rbForcedToPara, rbParaToForced now use button logic; no onClick here)

        // cbLead.onClick = requestPreview;
        // cbTrail.onClick = requestPreview;
        // cbMulti.onClick = requestPreview;

        cbSpaceToUnderscore.onClick = function () {
            if (cbSpaceToUnderscore.value) cbSpaceToHyphen.value = false;
            requestPreview();
        };
        cbSpaceToHyphen.onClick = function () {
            if (cbSpaceToHyphen.value) cbSpaceToUnderscore.value = false;
            requestPreview();
        };


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
        // ※行頭/行末の削除は「連続するスペース」のON/OFFに依存させない（常にまとめて削除）
        var spaces = " \u3000";
        var reLead = new RegExp("(^|[\\r\\n\\u0003])[" + spaces + "]+", "g");
        var reTrail = new RegExp("[" + spaces + "]+([\\r\\n\\u0003]|$)", "g");

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
        // --- ひらがな ↔ カタカナ ---
        function hiraToKata(s) {
            var out = "";
            for (var i = 0; i < s.length; i++) {
                var c = s.charCodeAt(i);
                // ぁ(3041)〜ゖ(3096), ゔ(3094), ゝ(309D)〜ゞ(309E)
                if ((c >= 0x3041 && c <= 0x3096) || c === 0x3094 || (c >= 0x309D && c <= 0x309E)) {
                    out += String.fromCharCode(c + 0x60);
                } else {
                    out += s.charAt(i);
                }
            }
            return out;
        }

        function kataToHira(s) {
            var out = "";
            for (var i = 0; i < s.length; i++) {
                var c = s.charCodeAt(i);
                // ァ(30A1)〜ヶ(30F6), ヴ(30F4), ヽ(30FD)〜ヾ(30FE)
                if ((c >= 0x30A1 && c <= 0x30F6) || c === 0x30F4 || (c >= 0x30FD && c <= 0x30FE)) {
                    out += String.fromCharCode(c - 0x60);
                } else {
                    out += s.charAt(i);
                }
            }
            return out;
        }

        // --- 全角カタカナ ↔ 半角カタカナ ---
        var __FULL_KATA_TO_HALF = {
            "。": "｡", "「": "｢", "」": "｣", "、": "､", "・": "･",
            "ァ": "ｧ", "ィ": "ｨ", "ゥ": "ｩ", "ェ": "ｪ", "ォ": "ｫ",
            "ャ": "ｬ", "ュ": "ｭ", "ョ": "ｮ", "ッ": "ｯ",
            "ー": "ｰ",
            "ア": "ｱ", "イ": "ｲ", "ウ": "ｳ", "エ": "ｴ", "オ": "ｵ",
            "カ": "ｶ", "キ": "ｷ", "ク": "ｸ", "ケ": "ｹ", "コ": "ｺ",
            "サ": "ｻ", "シ": "ｼ", "ス": "ｽ", "セ": "ｾ", "ソ": "ｿ",
            "タ": "ﾀ", "チ": "ﾁ", "ツ": "ﾂ", "テ": "ﾃ", "ト": "ﾄ",
            "ナ": "ﾅ", "ニ": "ﾆ", "ヌ": "ﾇ", "ネ": "ﾈ", "ノ": "ﾉ",
            "ハ": "ﾊ", "ヒ": "ﾋ", "フ": "ﾌ", "ヘ": "ﾍ", "ホ": "ﾎ",
            "マ": "ﾏ", "ミ": "ﾐ", "ム": "ﾑ", "メ": "ﾒ", "モ": "ﾓ",
            "ヤ": "ﾔ", "ユ": "ﾕ", "ヨ": "ﾖ",
            "ラ": "ﾗ", "リ": "ﾘ", "ル": "ﾙ", "レ": "ﾚ", "ロ": "ﾛ",
            "ワ": "ﾜ", "ヲ": "ｦ", "ン": "ﾝ",
            "ヴ": "ｳﾞ",
            "ガ": "ｶﾞ", "ギ": "ｷﾞ", "グ": "ｸﾞ", "ゲ": "ｹﾞ", "ゴ": "ｺﾞ",
            "ザ": "ｻﾞ", "ジ": "ｼﾞ", "ズ": "ｽﾞ", "ゼ": "ｾﾞ", "ゾ": "ｿﾞ",
            "ダ": "ﾀﾞ", "ヂ": "ﾁﾞ", "ヅ": "ﾂﾞ", "デ": "ﾃﾞ", "ド": "ﾄﾞ",
            "バ": "ﾊﾞ", "ビ": "ﾋﾞ", "ブ": "ﾌﾞ", "ベ": "ﾍﾞ", "ボ": "ﾎﾞ",
            "パ": "ﾊﾟ", "ピ": "ﾋﾟ", "プ": "ﾌﾟ", "ペ": "ﾍﾟ", "ポ": "ﾎﾟ",
            "ヷ": "ﾜﾞ", "ヺ": "ｦﾞ"
        };

        var __HALF_KATA_TO_FULL_SINGLE = {
            "｡": "。", "｢": "「", "｣": "」", "､": "、", "･": "・",
            "ｧ": "ァ", "ｨ": "ィ", "ｩ": "ゥ", "ｪ": "ェ", "ｫ": "ォ",
            "ｬ": "ャ", "ｭ": "ュ", "ｮ": "ョ", "ｯ": "ッ",
            "ｰ": "ー",
            "ｱ": "ア", "ｲ": "イ", "ｳ": "ウ", "ｴ": "エ", "ｵ": "オ",
            "ｶ": "カ", "ｷ": "キ", "ｸ": "ク", "ｹ": "ケ", "ｺ": "コ",
            "ｻ": "サ", "ｼ": "シ", "ｽ": "ス", "ｾ": "セ", "ｿ": "ソ",
            "ﾀ": "タ", "ﾁ": "チ", "ﾂ": "ツ", "ﾃ": "テ", "ﾄ": "ト",
            "ﾅ": "ナ", "ﾆ": "ニ", "ﾇ": "ヌ", "ﾈ": "ネ", "ﾉ": "ノ",
            "ﾊ": "ハ", "ﾋ": "ヒ", "ﾌ": "フ", "ﾍ": "ヘ", "ﾎ": "ホ",
            "ﾏ": "マ", "ﾐ": "ミ", "ﾑ": "ム", "ﾒ": "メ", "ﾓ": "モ",
            "ﾔ": "ヤ", "ﾕ": "ユ", "ﾖ": "ヨ",
            "ﾗ": "ラ", "ﾘ": "リ", "ﾙ": "ル", "ﾚ": "レ", "ﾛ": "ロ",
            "ﾜ": "ワ", "ｦ": "ヲ", "ﾝ": "ン",
            "ﾞ": "゛", "ﾟ": "゜"
        };

        var __HALF_KATA_TO_FULL_COMBO = {
            "ｳﾞ": "ヴ",
            "ｶﾞ": "ガ", "ｷﾞ": "ギ", "ｸﾞ": "グ", "ｹﾞ": "ゲ", "ｺﾞ": "ゴ",
            "ｻﾞ": "ザ", "ｼﾞ": "ジ", "ｽﾞ": "ズ", "ｾﾞ": "ゼ", "ｿﾞ": "ゾ",
            "ﾀﾞ": "ダ", "ﾁﾞ": "ヂ", "ﾂﾞ": "ヅ", "ﾃﾞ": "デ", "ﾄﾞ": "ド",
            "ﾊﾞ": "バ", "ﾋﾞ": "ビ", "ﾌﾞ": "ブ", "ﾍﾞ": "ベ", "ﾎﾞ": "ボ",
            "ﾊﾟ": "パ", "ﾋﾟ": "ピ", "ﾌﾟ": "プ", "ﾍﾟ": "ペ", "ﾎﾟ": "ポ",
            "ﾜﾞ": "ヷ", "ｦﾞ": "ヺ"
        };

        function kataToHalfwidth(s) {
            var out = "";
            for (var i = 0; i < s.length; i++) {
                var ch = s.charAt(i);
                out += (__FULL_KATA_TO_HALF.hasOwnProperty(ch) ? __FULL_KATA_TO_HALF[ch] : ch);
            }
            return out;
        }

        function halfwidthToKata(s) {
            var out = "";
            for (var i = 0; i < s.length; i++) {
                var ch = s.charAt(i);
                var next = (i + 1 < s.length) ? s.charAt(i + 1) : "";

                if (next === "ﾞ" || next === "ﾟ") {
                    var pair = ch + next;
                    if (__HALF_KATA_TO_FULL_COMBO.hasOwnProperty(pair)) {
                        out += __HALF_KATA_TO_FULL_COMBO[pair];
                        i++;
                        continue;
                    }
                }
                out += (__HALF_KATA_TO_FULL_SINGLE.hasOwnProperty(ch) ? __HALF_KATA_TO_FULL_SINGLE[ch] : ch);
            }
            return out;
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
            // ※行頭/行末の削除は「連続するスペース」のON/OFFに依存させない（常にまとめて削除）
            var spaces = " \u3000";
            var reLead = new RegExp("(^|[\\r\\n\\u0003])[" + spaces + "]+", "g");
            var reTrail = new RegExp("[" + spaces + "]+([\\r\\n\\u0003]|$)", "g");

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

            // --- any spaces (half/full) ---
            var reAnySpace = new RegExp("[" + spaces + "]+", "g");

            // --- hyphen / underscore ---
            var reHyphen = /-/g;
            var reUnderscore = /_/g;

            return {
                spaces: spaces,
                digits: digits,
                reLead: reLead,
                reTrail: reTrail,
                reAnySpace: reAnySpace,
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

            // 0) タブ→半角スペース（※［文字種］タブの「変換」ボタンでのみ実行）
            if (false && cbTabToSpace.value) {
                s = s.replace(rx.reTab, " ");
            }

            // ※スペース系（行頭/行末/連続/和文内/和欧内）は［スペース］タブ下部の［処理］ボタンでのみ実行する
            // 1) スペースの行頭・行末削除
            if (false && cbLead.value) {
                s = s.replace(rx.reLead, function (_m, p1) { return (p1 != null ? p1 : ""); });
            }
            if (false && cbTrail.value) {
                s = s.replace(rx.reTrail, function (_m, p1) { return (p1 != null ? p1 : ""); });
            }

            // 1.1) 連続するスペースの圧縮（行頭/行末は cbLead/cbTrail に従って保持）
            if (false && cbMulti.value) {
                var parts = s.split(/(\r|\u0003|\n)/);
                for (var i = 0; i < parts.length; i += 2) {
                    var line = parts[i] != null ? String(parts[i]) : "";

                    // 行頭/行末のスペースは触らない（インデント等の保持）
                    var mLead = line.match(/^[ \u3000]+/);
                    var mTrail = line.match(/[ \u3000]+$/);
                    var lead = mLead ? mLead[0] : "";
                    var trail = mTrail ? mTrail[0] : "";

                    var start = lead.length;
                    var end = line.length - trail.length;
                    if (end < start) { start = 0; end = line.length; lead = ""; trail = ""; }

                    var core = line.substring(start, end);
                    // 半角/全角スペースが2つ以上続く箇所を1つの半角スペースに
                    core = core.replace(/[ \u3000]{2,}/g, " ");

                    parts[i] = lead + core + trail;
                }
                s = parts.join("");
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

            // 5.4) 半角/全角スペース → アンダースコア／ハイフン
            // ※行頭/行末のスペースは触らない（インデント等の保持）
            if ((cbSpaceToUnderscore.value && !cbSpaceToHyphen.value) || (cbSpaceToHyphen.value && !cbSpaceToUnderscore.value)) {
                var repl = (cbSpaceToUnderscore.value ? "_" : "-");
                var partsS = s.split(/(\r|\u0003|\n)/);
                for (var si = 0; si < partsS.length; si += 2) {
                    var lineS = partsS[si] != null ? String(partsS[si]) : "";

                    var mLeadS = lineS.match(/^[ \u3000]+/);
                    var mTrailS = lineS.match(/[ \u3000]+$/);
                    var leadS = mLeadS ? mLeadS[0] : "";
                    var trailS = mTrailS ? mTrailS[0] : "";

                    var startS = leadS.length;
                    var endS = lineS.length - trailS.length;
                    if (endS < startS) { startS = 0; endS = lineS.length; leadS = ""; trailS = ""; }

                    var coreS = lineS.substring(startS, endS);
                    coreS = coreS.replace(rx.reAnySpace, repl);

                    partsS[si] = leadS + coreS + trailS;
                }
                s = partsS.join("");
            }

            // 5.5) 全角英数字を半角に（※［文字種］タブの「変換」ボタンでのみ実行）
            if (false && cbZenkakuAlnumToHankaku.value) {
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
        // alert("完了しました。");
    }) ();