#target illustrator
#targetengine "SelectionInspectorSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

SelectionInspector.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionInspector.jsx

### 概要：

- 選択中／全体のオブジェクト数をカウントし、2カラムのパレットで表示
- テキスト・配置画像・透明・グループ・パス・ガイドの統計を表示
- 「書き出し」でレポート（テキスト）を書き出し可能
- メモ（属性パネルのメモ欄）の閲覧・編集に対応
- 常駐パレット化：開いたまま選択を切り替え、［更新］で再集計
- DOM 集計・メモ書き込みはメインエンジンへ BridgeTalk 委譲（常駐パレットの DOM 切断を回避）

### 主な機能：

- オブジェクト数（選択／全体）・アートボード数を表示
- テキスト（ポイント／エリア／パス上）・文字数・段落数・強制改行数を表示
- 配置画像（リンク／埋め込み／リンク切れ）を表示
- 透明（不透明度<100、描画モード≠通常）を表示
- グループ／クリップグループを表示
- パス（オープン／クローズ／アンカー／ハンドル／複合パス／複合シェイプ）を表示
  ※ ガイド（guides=true）はパス統計から除外
- ガイド（ルーラー／アートボード／その他）を表示
- 「メモ」タブで選択オブジェクトのメモを編集・適用
- 「書き出し」ボタンで集計結果をテキストファイルとして保存

### 処理の流れ：

- 常駐パレットを表示（多重起動は自動で閉じてから再表示）
- ［更新］押下（または表示直後）にメインエンジンへ集計を委譲
- 戻り値（マーカー方式）を解析し、各パネル・メモタブを更新
- メモ適用・書き出しも同様に委譲／収集データから生成

### note：

https://note.com/dtp_tranist/n/nefcb1ce828ce

### 更新履歴：

- v1.0 (20250806) : 初期バージョン
- v1.1 (20250806) : 書き出し機能を追加
- v1.2 (20250807) : UI調整
- v1.3 (20260301) : パスのハンドル数を追加
- v1.4 (20260301) : 強制改行（ソフトリターン）数を追加
- v1.4.1 (20260302) : UI調整（OK→閉じる、キャンセル削除、書き出しを左へ移動）
- v1.5 (20260302) : ガイド集計（ルーラー/アートボード/その他）、ガイド除外のパス統計、パネル配置と書き出しレイアウト更新
- v1.5.1 (20260312) : グループ内のパス統計を再帰的にカウント
- v1.5.2 (20260312) : グループ内のテキスト統計を再帰的にカウント
- v1.5.3 (20260314) : ダイアログ位置をセッション中に記憶・復元
- v1.6 (20260316) : 配置画像の下にメモパネルを追加（属性パネルのメモ欄を表示）
- v1.6.1 (20260316) : ローカライズ整理（書き出し文言・セクション見出しをLABELSへ集約）
- v1.7.0 (20260702) : 常駐パレット化（#targetengine ＋ BridgeTalk 委譲）、更新ボタン・ステータス表示、ローカライズをカテゴリ構造＋L()へ再編

---

### Script Name:

SelectionInspector.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionInspector.jsx

### Overview:

- Count selection / document totals and show them in a two-column palette
- Show stats for text, images, transparency, groups, paths, and guides
- Export the stats as a plain text report
- View and edit object notes (Attributes panel note field)
- Persistent palette: keep it open, change selection, press Refresh to recount
- Delegate DOM counting / note writing to the main engine via BridgeTalk

### Main Features:

- Show object counts (Selection / All) and artboard count
- Text stats: point/area/path text, characters, paragraphs, forced line breaks
- Image stats: linked/embedded/broken links
- Transparency stats: opacity < 100, blend mode != Normal
- Group stats: groups and clipping groups
- Path stats: open/closed, anchor points, handles, compound paths, compound shapes
  * Guide paths (guides=true) are excluded from path stats
- Guide breakdown: ruler / artboard / other
- Edit and apply notes on the Notes tab
- Export button writes a plain text report to the desktop

### Process Flow:

- Show a persistent palette (existing one is closed first to prevent duplicates)
- On Refresh (or right after showing), delegate counting to the main engine
- Parse the marker-based result and update each panel and the Notes tab
- Note apply / export are delegated / generated from collected data

### Changelog:

- v1.0 (20250806): Initial version
- v1.1 (20250806): Added export feature
- v1.2 (20250807): UI adjustments
- v1.3 (20260301): Added handle count for paths
- v1.4 (20260301): Added forced line break (soft return) count
- v1.4.1 (20260302): UI tweaks (OK->Close, removed Cancel, moved Export to left)
- v1.5 (20260302): Added guide breakdown, excluded guides from path stats, updated layout
- v1.5.1 (20260312): Recursively count path stats inside groups
- v1.5.2 (20260312): Recursively count text stats inside groups
- v1.5.3 (20260314): Remember and restore dialog position during the current session
- v1.6 (20260316): Added Notes panel below Images
- v1.6.1 (20260316): Localization cleanup
- v1.7.0 (20260702): Palette conversion (#targetengine + BridgeTalk delegation), refresh button, status line, categorized localization with L()

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectionInspector";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* ============================================================
   言語判定・ローカライズ / Language & localization
   ============================================================ */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
var LABELS = {
    dialog: {
        title: { ja: "選択／全体オブジェクトのカウント", en: "Selection / All Objects Count" }
    },
    report: {
        title: { ja: "Selection Inspector Report", en: "Selection Inspector Report" },
        document: { ja: "ドキュメント:", en: "Document:" },
        date: { ja: "日付:", en: "Date:" },
        valueNote: {
            ja: "※ 値は『選択 / 全体』の形式です（アートボードのみ全体）",
            en: "Note: values are formatted as 'Selection / All' (Artboards is All only)"
        }
    },
    section: {
        basics: { ja: "基本", en: "Basics" },
        textFrames: { ja: "テキスト", en: "Text Frames" },
        charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
        images: { ja: "配置画像", en: "Images" },
        notes: { ja: "メモ", en: "Notes" },
        transparency: { ja: "透明", en: "Transparency" },
        groups: { ja: "グループ", en: "Groups" },
        paths: { ja: "パス", en: "Paths" },
        guides: { ja: "ガイド", en: "Guides" }
    },
    panel: {
        basics: { ja: "基本", en: "Basics" },
        texts: { ja: "テキスト", en: "Text Frames" },
        charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
        images: { ja: "配置画像", en: "Images" },
        memo: { ja: "メモ", en: "Notes" },
        group: { ja: "グループ", en: "Groups" },
        transparency: { ja: "透明", en: "Transparency" },
        path: { ja: "パス", en: "Paths" },
        guide: { ja: "ガイド", en: "Guides" }
    },
    row: {
        artboards: { ja: "アートボード：", en: "Artboards:" },
        objects: { ja: "オブジェクト：", en: "Objects:" },
        texts: { ja: "テキスト：", en: "Text Frames:" },
        pointText: { ja: "ポイント文字：", en: "Point Type:" },
        areaText: { ja: "エリア内文字：", en: "Area Type:" },
        pathText: { ja: "パス上文字：", en: "Path Text:" },
        chars: { ja: "文字数：", en: "Characters:" },
        paras: { ja: "段落数：", en: "Paragraphs:" },
        forcedBreaks: { ja: "強制改行：", en: "Line Breaks:" },
        linked: { ja: "リンク：", en: "Linked Images:" },
        embed: { ja: "埋め込み：", en: "Embedded Images:" },
        broken: { ja: "リンク切れ：", en: "Broken Links:" },
        group: { ja: "グループ：", en: "Groups:" },
        clipGroup: { ja: "クリップグループ：", en: "Clipping Groups:" },
        opacityLt100: { ja: "不透明度<100：", en: "Opacity < 100:" },
        blendNotNormal: { ja: "描画モード≠通常：", en: "Blend Mode != Normal:" },
        pathCount: { ja: "パス：", en: "Paths:" },
        openPath: { ja: "オープンパス：", en: "Open Paths:" },
        closedPath: { ja: "クローズパス：", en: "Closed Paths:" },
        anchors: { ja: "アンカーポイント：", en: "Anchor Points:" },
        handles: { ja: "ハンドル：", en: "Handles:" },
        compoundPath: { ja: "複合パス：", en: "Compound Paths:" },
        compoundShape: { ja: "複合シェイプ：", en: "Compound Shapes:" },
        rulerGuides: { ja: "ルーラーガイド：", en: "Ruler Guides:" },
        artboardGuides: { ja: "アートボードガイド：", en: "Artboard Guides:" },
        otherGuides: { ja: "その他のガイド：", en: "Other Guides:" }
    },
    button: {
        refresh: { ja: "更新", en: "Refresh" },
        exportPreset: { ja: "書き出し", en: "Export" },
        applyMemo: { ja: "適用", en: "Apply" }
    },
    tab: {
        info: { ja: "情報", en: "Info" },
        memo: { ja: "メモ", en: "Notes" }
    },
    memo: {
        multiple: { ja: "複数のメモがあります。\n「メモ」タブで確認", en: "Multiple notes found.\nSee the \"Notes\" tab." },
        none: { ja: "選択オブジェクトがありません", en: "No objects selected" }
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
        memoApplied: { ja: "メモを適用しました", en: "Note applied" },
        selChanged: { ja: "選択が変わりました。更新してください", en: "Selection changed. Please refresh." },
        exportedPrefix: { ja: "書き出しました: ", en: "Exported: " },
        exportFailOpen: { ja: "ファイルを開けませんでした", en: "Failed to open the file" }
    },
    hint: {
        shortcut: { ja: "⌥I: 情報  ⌥M: メモ  ⌘R: 更新", en: "⌥I: Info  ⌥M: Notes  ⌘R: Refresh" },
        refresh: { ja: "選択内容を再集計（⌘R）", en: "Recount selection (Cmd+R)" },
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

/* 書き出し用：末尾コロンを半角へ正規化 / Normalize trailing colon for export */
function LX(path) {
    return L(path).replace(/[：:]\s*$/, ":");
}

/* ============================================================
   定数 / Constants
   ============================================================ */
var LABEL_WIDTH_LEFT = 100;
var LABEL_WIDTH_RIGHT = 130;
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
function wkGetMaxArtboardSpan(doc) {
    var maxSpan = 0;
    try {
        var abs = Math.abs;
        var ab = doc.artboards;
        for (var i = 0; i < ab.length; i++) {
            var r = ab[i].artboardRect;
            var w = abs(r[2] - r[0]);
            var h = abs(r[1] - r[3]);
            if (w > maxSpan) { maxSpan = w; }
            if (h > maxSpan) { maxSpan = h; }
        }
    } catch (e) {}
    return maxSpan;
}

function wkIsArtboardGuide(pi, doc) {
    try {
        if (!pi || pi.typename !== "PathItem") { return false; }
        if (!pi.guides) { return false; }
        if (pi.closed) { return false; }
        if (!pi.pathPoints || pi.pathPoints.length !== 2) { return false; }
        var a0 = pi.pathPoints[0].anchor;
        var a1 = pi.pathPoints[1].anchor;
        var T = 0.01;
        var sameX = Math.abs(a0[0] - a1[0]) <= T;
        var sameY = Math.abs(a0[1] - a1[1]) <= T;
        if (!(sameX || sameY)) { return false; }
        var dx = a0[0] - a1[0];
        var dy = a0[1] - a1[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        var abs = Math.abs;
        var tol = 0.5;
        var ab = doc.artboards;
        for (var i = 0; i < ab.length; i++) {
            var r = ab[i].artboardRect;
            var w = abs(r[2] - r[0]);
            var h = abs(r[1] - r[3]);
            if (abs(len - w) <= tol) { return true; }
            if (abs(len - h) <= tol) { return true; }
        }
        return false;
    } catch (e) { return false; }
}

function wkIsRulerGuide(pi) {
    try {
        if (!pi || pi.typename !== "PathItem") { return false; }
        if (!pi.guides) { return false; }
        if (pi.closed) { return false; }
        if (!pi.pathPoints || pi.pathPoints.length !== 2) { return false; }
        var a0 = pi.pathPoints[0].anchor;
        var a1 = pi.pathPoints[1].anchor;
        var T = 0.01;
        var sameX = Math.abs(a0[0] - a1[0]) <= T;
        var sameY = Math.abs(a0[1] - a1[1]) <= T;
        if (!(sameX || sameY)) { return false; }
        var dx = a0[0] - a1[0];
        var dy = a0[1] - a1[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        var maxSpan = wkGetMaxArtboardSpan(app.activeDocument);
        if (!(maxSpan > 0)) { return true; }
        return (len > maxSpan);
    } catch (e) { return false; }
}

function wkCountGuides(pi, doc) {
    var r = { ruler: 0, artboard: 0, other: 0 };
    try {
        if (!pi || pi.typename !== "PathItem") { return r; }
        if (!pi.guides) { return r; }
        if (wkIsRulerGuide(pi)) { r.ruler = 1; }
        else if (wkIsArtboardGuide(pi, doc)) { r.artboard = 1; }
        else { r.other = 1; }
    } catch (e) {}
    return r;
}

function wkCountTransparency(it) {
    var r = { opacityLt100: 0, blendNotNormal: 0 };
    if (!it) { return r; }
    try { if (typeof it.opacity === "number" && it.opacity < 100) { r.opacityLt100 = 1; } } catch (e) {}
    try { if (it.blendingMode !== undefined && it.blendingMode !== BlendModes.NORMAL) { r.blendNotNormal = 1; } } catch (e2) {}
    return r;
}

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

function wkCountForcedBreaks(s) {
    if (!s || !s.length) { return 0; }
    var m = s.match(/[\n]/g);
    return m ? m.length : 0;
}

function wkCountTextStats(it, stats) {
    if (it.typename === "GroupItem") {
        for (var gi = 0; gi < it.pageItems.length; gi++) { wkCountTextStats(it.pageItems[gi], stats); }
    } else if (it.typename === "TextFrame") {
        stats.textCount++;
        try { stats.charCount += it.characters.length; } catch (e) {}
        try { stats.paraCount += it.paragraphs.length; } catch (e2) {}
        try { stats.forcedBreakCount += wkCountForcedBreaks(it.contents); } catch (e3) {}
        if (it.kind === TextType.POINTTEXT) { stats.pointText++; }
        else if (it.kind === TextType.AREATEXT) { stats.areaText++; }
        else if (it.kind === TextType.PATHTEXT) { stats.pathText++; }
    }
}

function wkSortSelection(sel) {
    var arr = [];
    for (var i = 0; i < sel.length; i++) { arr.push(sel[i]); }
    arr.sort(function (a, b) {
        var aTop = 0, bTop = 0, aLeft = 0, bLeft = 0;
        try { aTop = a.geometricBounds[1]; } catch (e) {}
        try { bTop = b.geometricBounds[1]; } catch (e2) {}
        try { aLeft = a.geometricBounds[0]; } catch (e3) {}
        try { bLeft = b.geometricBounds[0]; } catch (e4) {}
        if (aTop !== bTop) { return bTop - aTop; }
        return aLeft - bLeft;
    });
    return arr;
}

function wkCollect() {
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

    var allItems = doc.pageItems;

    var cpathSel = 0, cpathAll = 0, cshapeSel = 0, cshapeAll = 0;
    var opacitySel = 0, opacityAll = 0, blendSel = 0, blendAll = 0;
    var rulerSel = 0, rulerAll = 0, abguideSel = 0, abguideAll = 0, otherguideSel = 0, otherguideAll = 0;

    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "CompoundPathItem") { cpathSel++; }
        if (sel[i].typename === "PluginItem") {
            try { if (sel[i].name && sel[i].name.indexOf("Compound Shape") !== -1) { cshapeSel++; } } catch (e) {}
        }
        try { var t0 = wkCountTransparency(sel[i]); opacitySel += t0.opacityLt100; blendSel += t0.blendNotNormal; } catch (e2) {}
        try {
            if (sel[i].typename === "PathItem") {
                var g0 = wkCountGuides(sel[i], doc); rulerSel += g0.ruler; abguideSel += g0.artboard; otherguideSel += g0.other;
            } else if (sel[i].typename === "CompoundPathItem") {
                for (var gg = 0; gg < sel[i].pathItems.length; gg++) {
                    var g1 = wkCountGuides(sel[i].pathItems[gg], doc); rulerSel += g1.ruler; abguideSel += g1.artboard; otherguideSel += g1.other;
                }
            }
        } catch (e3) {}
    }

    var pathStatsSel = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
    for (var i2 = 0; i2 < sel.length; i2++) { wkCountPathStats(sel[i2], pathStatsSel); }

    for (var k = 0; k < allItems.length; k++) {
        var obj = allItems[k];
        if (obj.typename === "CompoundPathItem") { cpathAll++; }
        if (obj.typename === "PluginItem") {
            try { if (obj.name && obj.name.indexOf("Compound Shape") !== -1) { cshapeAll++; } } catch (e4) {}
        }
        try { var tA = wkCountTransparency(obj); opacityAll += tA.opacityLt100; blendAll += tA.blendNotNormal; } catch (e5) {}
        try {
            if (obj.typename === "PathItem") {
                var gA0 = wkCountGuides(obj, doc); rulerAll += gA0.ruler; abguideAll += gA0.artboard; otherguideAll += gA0.other;
            } else if (obj.typename === "CompoundPathItem") {
                for (var gg2 = 0; gg2 < obj.pathItems.length; gg2++) {
                    var gA1 = wkCountGuides(obj.pathItems[gg2], doc); rulerAll += gA1.ruler; abguideAll += gA1.artboard; otherguideAll += gA1.other;
                }
            }
        } catch (e6) {}
    }

    var pathStatsAll = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
    for (var k2 = 0; k2 < allItems.length; k2++) { wkCountPathStats(allItems[k2], pathStatsAll); }

    var textStatsSel = { textCount: 0, charCount: 0, paraCount: 0, forcedBreakCount: 0, pointText: 0, areaText: 0, pathText: 0 };
    for (var i3 = 0; i3 < sel.length; i3++) { wkCountTextStats(sel[i3], textStatsSel); }
    var textStatsAll = { textCount: 0, charCount: 0, paraCount: 0, forcedBreakCount: 0, pointText: 0, areaText: 0, pathText: 0 };
    for (var k3 = 0; k3 < allItems.length; k3++) { wkCountTextStats(allItems[k3], textStatsAll); }

    var linkedSel = 0, linkedAll = 0, embedSel = 0, embedAll = 0, brokenSel = 0, brokenAll = 0;
    for (var i4 = 0; i4 < sel.length; i4++) {
        if (sel[i4].typename === "PlacedItem") {
            if (sel[i4].embedded) { embedSel++; }
            else {
                linkedSel++;
                try { var f = sel[i4].file; if (!f || !f.exists) { brokenSel++; } } catch (e7) { brokenSel++; }
            }
        }
    }
    for (var k4 = 0; k4 < allItems.length; k4++) {
        var obj2 = allItems[k4];
        if (obj2.typename === "PlacedItem") {
            if (obj2.embedded) { embedAll++; }
            else {
                linkedAll++;
                try { var fA = obj2.file; if (!fA || !fA.exists) { brokenAll++; } } catch (e8) { brokenAll++; }
            }
        }
    }

    var sorted = wkSortSelection(sel);
    var memoParts = [];
    for (var mi = 0; mi < sorted.length; mi++) {
        var nt = "";
        try { nt = sorted[mi].note || ""; } catch (e9) { nt = ""; }
        memoParts.push(encodeURIComponent(nt));
    }

    var groupSel = 0, groupAll = 0, clipSel = 0, clipAll = 0;
    for (var i5 = 0; i5 < sel.length; i5++) {
        if (sel[i5].typename === "GroupItem") { groupSel++; if (sel[i5].clipped) { clipSel++; } }
    }
    for (var k5 = 0; k5 < allItems.length; k5++) {
        if (allItems[k5].typename === "GroupItem") { groupAll++; if (allItems[k5].clipped) { clipAll++; } }
    }

    var artboards = doc.artboards.length;
    var docName = "";
    try { docName = doc.name; } catch (e10) { docName = ""; }

    var out = [];
    out.push("selCount=" + selCount);
    out.push("allCount=" + allCount);
    out.push("artboards=" + artboards);
    out.push("docName=" + encodeURIComponent(docName));
    out.push("textSel=" + textStatsSel.textCount);
    out.push("textAll=" + textStatsAll.textCount);
    out.push("pointSel=" + textStatsSel.pointText);
    out.push("pointAll=" + textStatsAll.pointText);
    out.push("areaSel=" + textStatsSel.areaText);
    out.push("areaAll=" + textStatsAll.areaText);
    out.push("pathSel=" + textStatsSel.pathText);
    out.push("pathAll=" + textStatsAll.pathText);
    out.push("charSel=" + textStatsSel.charCount);
    out.push("charAll=" + textStatsAll.charCount);
    out.push("paraSel=" + textStatsSel.paraCount);
    out.push("paraAll=" + textStatsAll.paraCount);
    out.push("fbSel=" + textStatsSel.forcedBreakCount);
    out.push("fbAll=" + textStatsAll.forcedBreakCount);
    out.push("linkedSel=" + linkedSel);
    out.push("linkedAll=" + linkedAll);
    out.push("embedSel=" + embedSel);
    out.push("embedAll=" + embedAll);
    out.push("brokenSel=" + brokenSel);
    out.push("brokenAll=" + brokenAll);
    out.push("groupSel=" + groupSel);
    out.push("groupAll=" + groupAll);
    out.push("clipSel=" + clipSel);
    out.push("clipAll=" + clipAll);
    out.push("opacitySel=" + opacitySel);
    out.push("opacityAll=" + opacityAll);
    out.push("blendSel=" + blendSel);
    out.push("blendAll=" + blendAll);
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
    out.push("rulerSel=" + rulerSel);
    out.push("rulerAll=" + rulerAll);
    out.push("abguideSel=" + abguideSel);
    out.push("abguideAll=" + abguideAll);
    out.push("otherguideSel=" + otherguideSel);
    out.push("otherguideAll=" + otherguideAll);

    return "OK|" + out.join("|") + "|MEMO|" + selCount + "|" + memoParts.join("|");
}

function wkApplyMemo(index, enc) {
    if (app.documents.length === 0) { return "NODOC"; }
    var sel = app.activeDocument.selection;
    if (!sel) { sel = []; }
    if (index < 0 || index >= sel.length) { return "IDX"; }
    var sorted = wkSortSelection(sel);
    try { sorted[index].note = decodeURIComponent(enc); } catch (e) { return "ERR:" + e; }
    try { app.redraw(); } catch (e2) {}
    return "OK";
}

/* worker 関数は全登録（追加漏れ防止） / Register every worker function */
var WORKER_FUNCS = [
    wkGetMaxArtboardSpan,
    wkIsArtboardGuide,
    wkIsRulerGuide,
    wkCountGuides,
    wkCountTransparency,
    wkIsGuidePath,
    wkCountHandles,
    wkCountPathStats,
    wkCountForcedBreaks,
    wkCountTextStats,
    wkSortSelection,
    wkCollect,
    wkApplyMemo
];

/* ============================================================
   BridgeTalk 委譲 / Delegation to the main engine
   ============================================================ */
var isBusy = false;

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

/* 戻り値（OK|key=value|...|MEMO|count|note...）を解析 / Parse collect result */
function parseCollect(resp) {
    if (!resp || resp.indexOf("OK|") !== 0) return null;
    var raw = resp.substring(3);
    var idx = raw.indexOf("|MEMO|");
    if (idx < 0) return null;
    var statPart = raw.substring(0, idx);
    var memoPart = raw.substring(idx + 6);

    var map = {};
    var pairs = statPart.split("|");
    for (var i = 0; i < pairs.length; i++) {
        var eq = pairs[i].indexOf("=");
        if (eq > 0) { map[pairs[i].substring(0, eq)] = pairs[i].substring(eq + 1); }
    }
    if (map.docName != null) { try { map.docName = decodeURIComponent(map.docName); } catch (e) {} }

    var mm = memoPart.split("|");
    var count = parseInt(mm[0], 10) || 0;
    var memoList = [];
    for (var k = 1; k <= count && k < mm.length; k++) {
        var s = mm[k];
        try { s = decodeURIComponent(s); } catch (e2) {}
        memoList.push(s);
    }
    return { map: map, memoList: memoList };
}

/* ============================================================
   状態保持（常駐エンジン） / Session state (resident engine)
   ============================================================ */
if (!$.global.__selectionInspectorState) {
    $.global.__selectionInspectorState = { location: null };
}

/* ============================================================
   パレット構築 / Build palette
   ============================================================ */
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

function buildPalette() {
    var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    win.orientation = "column";
    win.alignChildren = "center";
    win.margins = [15, 10, 15, 15];

    var lastData = null;
    var memoFields = [];

    /* 表示切り替え（ラジオボタン＋stack）/ View switcher (radio buttons + stack) */
    var switchRow = win.add("group");
    switchRow.orientation = "row";
    switchRow.alignChildren = ["center", "center"];
    switchRow.helpTip = L('hint.shortcut');

    var rbInfo = switchRow.add("radiobutton", undefined, L('tab.info'));
    var rbMemo = switchRow.add("radiobutton", undefined, L('tab.memo'));
    rbInfo.helpTip = L('hint.shortcut');
    rbMemo.helpTip = L('hint.shortcut');
    rbInfo.value = true;

    var stackWrap = win.add("group");
    stackWrap.orientation = "stack";
    stackWrap.alignChildren = ["fill", "fill"];

    var tab1 = stackWrap.add("group");
    tab1.orientation = "column";
    tab1.alignChildren = ["fill", "top"];
    tab1.margins = [10, 15, 10, 10];

    var tab2 = stackWrap.add("group");
    tab2.orientation = "column";
    tab2.alignChildren = ["fill", "top"];
    tab2.margins = [10, 15, 10, 10];
    tab2.visible = false;

    /* --- 情報タブ（2カラム） / Info tab (two columns) --- */
    var twoColGroup = tab1.add("group");
    twoColGroup.orientation = "row";
    twoColGroup.alignChildren = ["fill", "top"];

    var leftCol = twoColGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];

    var rightCol = twoColGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];

    function addPanel(col, titlePath) {
        var p = col.add("panel", undefined, L(titlePath));
        p.orientation = "column";
        p.alignChildren = ["fill", "top"];
        p.margins = PANEL_MARGINS;
        return p;
    }

    var v = {};

    /* 左カラム / Left column */
    var panelBasics = addPanel(leftCol, 'panel.basics');
    v.artboards = addStatRow(panelBasics, L('row.artboards'), LABEL_WIDTH_LEFT);
    v.objects = addStatRow(panelBasics, L('row.objects'), LABEL_WIDTH_LEFT);

    var panelText = addPanel(leftCol, 'panel.texts');
    v.texts = addStatRow(panelText, L('row.texts'), LABEL_WIDTH_LEFT);
    v.pointText = addStatRow(panelText, L('row.pointText'), LABEL_WIDTH_LEFT);
    v.areaText = addStatRow(panelText, L('row.areaText'), LABEL_WIDTH_LEFT);
    v.pathText = addStatRow(panelText, L('row.pathText'), LABEL_WIDTH_LEFT);

    var panelCharPara = addPanel(leftCol, 'panel.charPara');
    v.chars = addStatRow(panelCharPara, L('row.chars'), LABEL_WIDTH_LEFT);
    v.paras = addStatRow(panelCharPara, L('row.paras'), LABEL_WIDTH_LEFT);
    v.forcedBreaks = addStatRow(panelCharPara, L('row.forcedBreaks'), LABEL_WIDTH_LEFT);

    var panelImage = addPanel(leftCol, 'panel.images');
    v.linked = addStatRow(panelImage, L('row.linked'), LABEL_WIDTH_LEFT);
    v.embed = addStatRow(panelImage, L('row.embed'), LABEL_WIDTH_LEFT);
    v.broken = addStatRow(panelImage, L('row.broken'), LABEL_WIDTH_LEFT);

    var panelMemo = addPanel(leftCol, 'panel.memo');
    var memoReadOnly = panelMemo.add("statictext", undefined, "", { multiline: true });
    memoReadOnly.preferredSize = [160, 50];

    /* 右カラム / Right column */
    var panelGroup = addPanel(rightCol, 'panel.group');
    v.group = addStatRow(panelGroup, L('row.group'), LABEL_WIDTH_RIGHT);
    v.clipGroup = addStatRow(panelGroup, L('row.clipGroup'), LABEL_WIDTH_RIGHT);

    var panelTransparency = addPanel(rightCol, 'panel.transparency');
    v.opacityLt100 = addStatRow(panelTransparency, L('row.opacityLt100'), LABEL_WIDTH_RIGHT);
    v.blendNotNormal = addStatRow(panelTransparency, L('row.blendNotNormal'), LABEL_WIDTH_RIGHT);

    var panelPath = addPanel(rightCol, 'panel.path');
    v.pathCount = addStatRow(panelPath, L('row.pathCount'), LABEL_WIDTH_RIGHT);
    v.openPath = addStatRow(panelPath, L('row.openPath'), LABEL_WIDTH_RIGHT);
    v.closedPath = addStatRow(panelPath, L('row.closedPath'), LABEL_WIDTH_RIGHT);
    v.anchors = addStatRow(panelPath, L('row.anchors'), LABEL_WIDTH_RIGHT);
    v.handles = addStatRow(panelPath, L('row.handles'), LABEL_WIDTH_RIGHT);
    v.compoundPath = addStatRow(panelPath, L('row.compoundPath'), LABEL_WIDTH_RIGHT);
    v.compoundShape = addStatRow(panelPath, L('row.compoundShape'), LABEL_WIDTH_RIGHT);

    var panelGuide = addPanel(rightCol, 'panel.guide');
    v.rulerGuides = addStatRow(panelGuide, L('row.rulerGuides'), LABEL_WIDTH_RIGHT);
    v.artboardGuides = addStatRow(panelGuide, L('row.artboardGuides'), LABEL_WIDTH_RIGHT);
    v.otherGuides = addStatRow(panelGuide, L('row.otherGuides'), LABEL_WIDTH_RIGHT);

    /* ステータス / Status line */
    var statusText = win.add("statictext", undefined, L('status.ready'));
    statusText.alignment = ["fill", "bottom"];

    function setStatus(msg) { try { statusText.text = msg; } catch (e) {} }

    function relayout() {
        try { if (stackWrap.layout) { stackWrap.layout.layout(true); } } catch (e) {}
        try { if (win.layout) { win.layout.layout(true); } } catch (e2) {}
    }

    function switchView(mode) {
        var showInfo = (mode === "info");
        rbInfo.value = showInfo;
        rbMemo.value = !showInfo;
        tab1.visible = showInfo;
        tab2.visible = !showInfo;
        relayout();
    }

    /* 値の一括更新 / Apply all values */
    function applyValues(m) {
        v.artboards.text = m.artboards;
        v.objects.text = m.selCount + " / " + m.allCount;
        v.texts.text = m.textSel + " / " + m.textAll;
        v.pointText.text = m.pointSel + " / " + m.pointAll;
        v.areaText.text = m.areaSel + " / " + m.areaAll;
        v.pathText.text = m.pathSel + " / " + m.pathAll;
        v.chars.text = m.charSel + " / " + m.charAll;
        v.paras.text = m.paraSel + " / " + m.paraAll;
        v.forcedBreaks.text = m.fbSel + " / " + m.fbAll;
        v.linked.text = m.linkedSel + " / " + m.linkedAll;
        v.embed.text = m.embedSel + " / " + m.embedAll;
        v.broken.text = m.brokenSel + " / " + m.brokenAll;
        v.group.text = m.groupSel + " / " + m.groupAll;
        v.clipGroup.text = m.clipSel + " / " + m.clipAll;
        v.opacityLt100.text = m.opacitySel + " / " + m.opacityAll;
        v.blendNotNormal.text = m.blendSel + " / " + m.blendAll;
        v.pathCount.text = m.pathCountSel + " / " + m.pathCountAll;
        v.openPath.text = m.openSel + " / " + m.openAll;
        v.closedPath.text = m.closedSel + " / " + m.closedAll;
        v.anchors.text = m.anchorSel + " / " + m.anchorAll;
        v.handles.text = m.handleSel + " / " + m.handleAll;
        v.compoundPath.text = m.cpathSel + " / " + m.cpathAll;
        v.compoundShape.text = m.cshapeSel + " / " + m.cshapeAll;
        v.rulerGuides.text = m.rulerSel + " / " + m.rulerAll;
        v.artboardGuides.text = m.abguideSel + " / " + m.abguideAll;
        v.otherGuides.text = m.otherguideSel + " / " + m.otherguideAll;
    }

    function clearValues() {
        for (var kk in v) { if (v.hasOwnProperty(kk)) { try { v[kk].text = "-"; } catch (e) {} } }
    }

    function updateMemoReadonly(memoList) {
        var nonEmpty = [];
        for (var i = 0; i < memoList.length; i++) {
            if (memoList[i] && memoList[i] !== "") { nonEmpty.push(memoList[i]); }
        }
        var text = (nonEmpty.length === 1) ? nonEmpty[0] : (nonEmpty.length > 1 ? L('memo.multiple') : "");
        try { memoReadOnly.text = text; } catch (e) {}
    }

    /* メモタブを再構築 / Rebuild the Notes tab */
    function rebuildMemo(memoList) {
        while (tab2.children.length > 0) {
            try { tab2.remove(tab2.children[0]); } catch (e) { break; }
        }
        memoFields = [];

        if (!memoList || memoList.length === 0) {
            tab2.add("statictext", undefined, L('memo.none'));
            relayout();
            return;
        }

        for (var i = 0; i < memoList.length; i++) {
            (function (idx, noteText) {
                var row = tab2.add("group");
                row.orientation = "row";
                row.alignChildren = ["fill", "center"];
                var field = row.add("edittext", undefined, noteText, { multiline: true });
                field.preferredSize = [340, 44];
                memoFields.push(field);
                var applyBtn = row.add("button", undefined, L('button.applyMemo'));
                applyBtn.onClick = function () { applyMemo(idx, field.text); };
            })(i, memoList[i]);
        }
        relayout();
    }

    /* 再集計 / Recount */
    function refresh() {
        setStatus(L('status.busy'));
        var resp = callMainEngine("wkCollect()");

        if (resp === "ERR:BUSY") { setStatus(L('status.busy')); return; }
        if (resp === null || resp === "ERR:TIMEOUT") { setStatus(L('status.timeout')); return; }
        if (resp === "NODOC") { setStatus(L('status.noDoc')); clearValues(); updateMemoReadonly([]); rebuildMemo([]); return; }
        if (resp.indexOf("ERR:") === 0) { setStatus(L('status.error') + ": " + resp.substring(4)); return; }

        var data = parseCollect(resp);
        if (!data) { setStatus(L('status.error')); return; }

        lastData = data;
        applyValues(data.map);
        updateMemoReadonly(data.memoList);
        rebuildMemo(data.memoList);

        var selN = parseInt(data.map.selCount, 10) || 0;
        if (selN > 0) {
            setStatus(L('status.selectedPrefix') + selN + L('status.selectedSuffix'));
        } else {
            setStatus(L('status.wholeDoc'));
        }
    }

    /* メモ適用 / Apply a note */
    function applyMemo(idx, text) {
        setStatus(L('status.busy'));
        var resp = callMainEngine("wkApplyMemo(" + idx + ",\"" + encodeURIComponent(text) + "\")");
        if (resp === "OK") { setStatus(L('status.memoApplied')); refresh(); }
        else if (resp === "NODOC") { setStatus(L('status.noDoc')); }
        else if (resp === "IDX") { setStatus(L('status.selChanged')); }
        else { setStatus(L('status.error') + ": " + resp); }
    }

    /* レポート書き出し（収集データからパレット側で生成） / Export report */
    function exportReport() {
        setStatus(L('status.busy'));
        var resp = callMainEngine("wkCollect()");
        if (resp === "NODOC") { setStatus(L('status.noDoc')); return; }
        if (resp === null || resp === "ERR:TIMEOUT") { setStatus(L('status.timeout')); return; }
        if (typeof resp === "string" && resp.indexOf("ERR:") === 0) { setStatus(L('status.error') + ": " + resp.substring(4)); return; }

        var data = parseCollect(resp);
        if (!data) { setStatus(L('status.error')); return; }
        var m = data.map;

        try {
            var fullName = m.docName || "";
            var baseName = fullName.replace(/\.[^\.]+$/, "");
            var today = new Date();
            var yyyy = today.getFullYear();
            var mm = ("0" + (today.getMonth() + 1)).slice(-2);
            var dd = ("0" + today.getDate()).slice(-2);
            var dateStr = yyyy + mm + dd;

            var path = Folder.desktop + "/count-" + baseName + "-" + dateStr + ".txt";
            var file = new File(path);

            function wPair(path2, selVal, allVal) { file.writeln(LX(path2) + " " + selVal + " / " + allVal); }
            function wSingle(path2, val) { file.writeln(LX(path2) + " " + val); }
            function wSection(path2) { file.writeln(""); file.writeln(L(path2)); }

            if (file.open("w")) {
                file.writeln(L('report.title'));
                file.writeln(L('report.document') + " " + fullName);
                file.writeln(L('report.date') + " " + yyyy + "-" + mm + "-" + dd);
                file.writeln("");
                file.writeln(L('report.valueNote'));

                wSection('section.basics');
                wSingle('row.artboards', m.artboards);
                wPair('row.objects', m.selCount, m.allCount);

                wSection('section.textFrames');
                wPair('row.texts', m.textSel, m.textAll);
                wPair('row.pointText', m.pointSel, m.pointAll);
                wPair('row.areaText', m.areaSel, m.areaAll);
                wPair('row.pathText', m.pathSel, m.pathAll);

                wSection('section.charPara');
                wPair('row.chars', m.charSel, m.charAll);
                wPair('row.paras', m.paraSel, m.paraAll);
                wPair('row.forcedBreaks', m.fbSel, m.fbAll);

                wSection('section.images');
                wPair('row.linked', m.linkedSel, m.linkedAll);
                wPair('row.embed', m.embedSel, m.embedAll);
                wPair('row.broken', m.brokenSel, m.brokenAll);

                wSection('section.notes');
                var nonEmpty = [];
                for (var ni = 0; ni < data.memoList.length; ni++) {
                    if (data.memoList[ni] && data.memoList[ni] !== "") { nonEmpty.push(data.memoList[ni]); }
                }
                if (nonEmpty.length > 0) { file.writeln(nonEmpty.join("\n")); }

                wSection('section.transparency');
                wPair('row.opacityLt100', m.opacitySel, m.opacityAll);
                wPair('row.blendNotNormal', m.blendSel, m.blendAll);

                wSection('section.groups');
                wPair('row.group', m.groupSel, m.groupAll);
                wPair('row.clipGroup', m.clipSel, m.clipAll);

                wSection('section.paths');
                wPair('row.pathCount', m.pathCountSel, m.pathCountAll);
                wPair('row.openPath', m.openSel, m.openAll);
                wPair('row.closedPath', m.closedSel, m.closedAll);
                wPair('row.anchors', m.anchorSel, m.anchorAll);
                wPair('row.handles', m.handleSel, m.handleAll);
                wPair('row.compoundPath', m.cpathSel, m.cpathAll);
                wPair('row.compoundShape', m.cshapeSel, m.cshapeAll);

                wSection('section.guides');
                wPair('row.rulerGuides', m.rulerSel, m.rulerAll);
                wPair('row.artboardGuides', m.abguideSel, m.abguideAll);
                wPair('row.otherGuides', m.otherguideSel, m.otherguideAll);

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

    rbInfo.onClick = function () { switchView("info"); };
    rbMemo.onClick = function () {
        switchView("memo");
        try { if (memoFields.length > 0) { memoFields[0].active = true; } } catch (e) {}
    };

    /* キー操作 / Key handling
       Esc: 閉じる / close
       ⌥I / ⌥M: タブ切替 / switch tabs
       ⌘R: 更新（Enter はメモ編集と衝突するため使わない）/ Cmd+R refresh */
    win.addEventListener("keydown", function (ev) {
        var key = "";
        try { key = ev && ev.keyName ? String(ev.keyName).toUpperCase() : ""; } catch (e) { key = ""; }

        if (key === "ESCAPE") {
            try { win.close(); } catch (e1) {}
        } else if (ev && ev.altKey && key === "I") {
            switchView("info");
            try { if (ev.preventDefault) ev.preventDefault(); } catch (e2) {}
        } else if (ev && ev.altKey && key === "M") {
            switchView("memo");
            try { if (memoFields.length > 0) { memoFields[0].active = true; } } catch (e3) {}
            try { if (ev.preventDefault) ev.preventDefault(); } catch (e4) {}
        } else if (ev && ev.metaKey && key === "R") {
            refresh();
            try { if (ev.preventDefault) ev.preventDefault(); } catch (e5) {}
        }
    });

    try { win.opacity = PALETTE_OPACITY; } catch (e) {}

    switchView("info");

    /* 表示直後に一度集計 / Count once right after showing */
    win.onShow = function () {
        refresh();
    };

    return win;
}

/* ============================================================
   位置の記憶・復元 / Remember & restore location
   ============================================================ */
function restoreLocation(win) {
    try {
        var loc = $.global.__selectionInspectorState.location;
        if (loc && loc.length === 2) {
            win.location = [loc[0], loc[1]];
        } else {
            win.center();
        }
    } catch (e) {
        win.center();
    }
}

function rememberLocation(win) {
    try {
        if (win.location && win.location.length === 2) {
            $.global.__selectionInspectorState.location = [win.location[0], win.location[1]];
        }
    } catch (e) {}
}

/* ============================================================
   起動 / Entry point
   ============================================================ */
function showPalette() {
    /* 多重起動防止：既存パレットがあれば閉じる / Prevent duplicates */
    if ($.global.__SelectionInspectorPalette) {
        try { $.global.__SelectionInspectorPalette.close(); } catch (e) {}
        $.global.__SelectionInspectorPalette = null;
    }

    var win = buildPalette();

    /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
    $.global.__SelectionInspectorPalette = win;
    win.onClose = function () {
        rememberLocation(win);
        try { app.redraw(); } catch (e) {}
        try { $.global.__SelectionInspectorPalette = null; } catch (e2) {}
    };

    restoreLocation(win);
    win.show();
}

showPalette();
