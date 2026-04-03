#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* =========================================
 * 書式を保ったままテキストを1文字ごとに分割（高精度）
 *
 * 概要:
 *  - 選択した TextFrame を 1文字ごとの TextFrame に分割します。
 *  - 最優先で「計算用にアウトライン化した各文字グループの座標（bounds）」を利用し、
 *    文字の見た目位置を可能な限り正確に再現します。
 *  - アウトライン化したオブジェクトは計算にのみ使用し、処理後に削除します。
 *  - もしアウトライン化で文字グループの取得ができない等の場合は、従来の
 *    「仮TextFrameで幅を積算して配置」方式にフォールバックします。
 *
 * 作成日: 2026-02-16
 * 更新日: 2026-02-16
 * ========================================= */

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var optKeepSpaces = false;
var optGroupMode = "none"; // none | line | all
var optSplitMode = "charVisual"; // charVisual | charEven | word | line
var optConvertToAreaText = false;
var optMergeAreaText = false;
var optKeepStyle = true;

var LABELS = {
    dialogTitle: {
        ja: "テキスト分割",
        en: "Text Split"
    },
    uiDesc: {
        ja: "選択したテキストを1文字ごとに分割します（スペースは無視）。",
        en: "Split selected text into single characters (spaces are ignored)."
    },
    chkKeepSpaces: {
        ja: "スペースを残す",
        en: "Keep spaces"
    },
    chkKeepStyle: {
        ja: "書式を保持",
        en: "Keep style"
    },
    chkConvertToAreaText: {
        ja: "エリア内文字に変換",
        en: "Convert to area type"
    },
    chkMergeAreaText: {
        ja: "エリア内文字を連結",
        en: "Merge area text"
    },
    pnlGroup: {
        ja: "グループ化",
        en: "Grouping"
    },
    rbGroupNone: {
        ja: "グループ化しない",
        en: "Do not group"
    },
    rbGroupLine: {
        ja: "行ごとにグループ化",
        en: "Group each line"
    },
    rbGroupAll: {
        ja: "全体をグループ化",
        en: "Group all"
    },
    btnOK: {
        ja: "OK",
        en: "OK"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    msgNoDoc: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    msgNoSelection: {
        ja: "テキストを選択してください。",
        en: "Please select text."
    },
    pnlSplitMode: {
        ja: "分割方法",
        en: "Split mode"
    },
    rbSplitCharVisual: {
        ja: "文字",
        en: "Character"
    },
    rbSplitWord: {
        ja: "単語",
        en: "Word"
    },
    rbSplitLine: {
        ja: "行",
        en: "Line"
    },
};

function L(key) {
    try {
        var o = LABELS[key];
        if (!o) return key;
        return o[lang] || o.ja || key;
    } catch (_) {
        return key;
    }
}

var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);

dlg.orientation = 'column';
dlg.alignChildren = ['fill', 'top'];
dlg.margins = 15;

// --- UI panel order: 分割方法 (SplitMode) → オプション (Opt) → グループ化 (Group) ---

var pnlSplitMode = dlg.add('panel', undefined, L('pnlSplitMode'));
pnlSplitMode.orientation = 'column';
pnlSplitMode.alignChildren = ['left', 'top'];
pnlSplitMode.margins = [15, 20, 15, 10];

var rbSplitCharVisual = pnlSplitMode.add('radiobutton', undefined, L('rbSplitCharVisual'));
var rbSplitWord = pnlSplitMode.add('radiobutton', undefined, L('rbSplitWord'));
var rbSplitLine = pnlSplitMode.add('radiobutton', undefined, L('rbSplitLine'));
rbSplitCharVisual.value = true; // default

var pnlOpt = dlg.add('panel', undefined, 'オプション');
pnlOpt.orientation = 'column';
pnlOpt.alignChildren = ['left', 'top'];
pnlOpt.margins = [15, 20, 15, 10];

var chkKeepStyle = pnlOpt.add('checkbox', undefined, L('chkKeepStyle'));
chkKeepStyle.value = true;
var chkKeepSpaces = pnlOpt.add('checkbox', undefined, L('chkKeepSpaces'));
chkKeepSpaces.value = false;
var chkConvertToAreaText = pnlOpt.add('checkbox', undefined, L('chkConvertToAreaText'));
chkConvertToAreaText.value = false;
var chkMergeAreaText = pnlOpt.add('checkbox', undefined, L('chkMergeAreaText'));
chkMergeAreaText.value = false;

// 「単語」「行」のときはエリア系オプションをディム（変換/連結とも）
function syncAreaOptionsEnabled() {
    try {
        var isWordOrLine = false;
        try {
            isWordOrLine = !!rbSplitWord.value || !!rbSplitLine.value;
        } catch (_) { isWordOrLine = false; }

        // split mode が単語/行ならディム
        chkConvertToAreaText.enabled = !isWordOrLine;
        if (!chkConvertToAreaText.enabled) chkConvertToAreaText.value = false;

        // 連結は「変換がON」かつ「単語/行ではない」時だけ有効
        chkMergeAreaText.enabled = (!isWordOrLine) && (!!chkConvertToAreaText.value);
        if (!chkMergeAreaText.enabled) chkMergeAreaText.value = false;
    } catch (_) { }
}

try {
    chkConvertToAreaText.onClick = function () {
        syncAreaOptionsEnabled();
    };
} catch (_) { }

// 分割方法の選択変更でも同期
try {
    rbSplitCharVisual.onClick = function () { syncAreaOptionsEnabled(); };
    rbSplitWord.onClick = function () { syncAreaOptionsEnabled(); };
    rbSplitLine.onClick = function () { syncAreaOptionsEnabled(); };
} catch (_) { }

// 初期状態
syncAreaOptionsEnabled();

var pnlGroup = dlg.add('panel', undefined, L('pnlGroup'));
pnlGroup.orientation = 'column';
pnlGroup.alignChildren = ['left', 'top'];
pnlGroup.margins = [15, 20, 15, 10];

var rbGroupNone = pnlGroup.add('radiobutton', undefined, L('rbGroupNone'));
var rbGroupLine = pnlGroup.add('radiobutton', undefined, L('rbGroupLine'));
var rbGroupAll = pnlGroup.add('radiobutton', undefined, L('rbGroupAll'));
rbGroupNone.value = true; // default


var gBtns = dlg.add('group');
gBtns.alignment = 'right';
var btnCancel = gBtns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
var btnOK = gBtns.add('button', undefined, L('btnOK'), { name: 'ok' });

btnOK.onClick = function () {
    try { optKeepSpaces = !!chkKeepSpaces.value; } catch (_) { optKeepSpaces = false; }
    try { optKeepStyle = !!chkKeepStyle.value; } catch (_) { optKeepStyle = true; }
    try { optConvertToAreaText = !!chkConvertToAreaText.value; } catch (_) { optConvertToAreaText = false; }
    try { optMergeAreaText = (!!chkConvertToAreaText.value) && (!!chkMergeAreaText.value); } catch (_) { optMergeAreaText = false; }
    try {
        optGroupMode = rbGroupNone.value ? "none" : (rbGroupLine.value ? "line" : "all");
    } catch (_) {
        optGroupMode = "none";
    }
    try {
        optSplitMode = rbSplitCharVisual.value ? "charVisual" : (rbSplitWord.value ? "word" : "line");
    } catch (_) {
        optSplitMode = "charVisual";
    }
    try { main(); } catch (_) { }
    try { dlg.close(1); } catch (_) { }
};
btnCancel.onClick = function () {
    try { dlg.close(0); } catch (_) { }
};

dlg.show();

function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length === 0) return;

    /* TextFrame のみ抽出 / Collect TextFrames only */
    var targets = [];
    for (var i = 0; i < sel.length; i++) {
        try {
            if (sel[i] && sel[i].typename === "TextFrame") targets.push(sel[i]);
        } catch (_) { }
    }
    if (targets.length === 0) return;

    for (var k = 0; k < targets.length; k++) {
        // 「書式を保持」OFF: 分割前に書式を整理（先頭文字のフォント情報のみ保持）
        try {
            if (optSplitMode === "charVisual" && !optKeepStyle) {
                stripStyleKeepFirstFont(targets[k]);
            }
        } catch (_) { }
        try {
            splitTextFrameHighPrecision(targets[k]);
        } catch (_) { }
    }
}

/* =========================================
 * Style reset / 書式削除（先頭文字のフォント情報だけ保持）
 * ========================================= */

function stripStyleKeepFirstFont(textFrame) {
    if (!textFrame || textFrame.typename !== "TextFrame") return;

    var tr = null;
    try { tr = textFrame.textRange; } catch (_) { tr = null; }
    if (!tr) return;

    var chars = null;
    try { chars = tr.characters; } catch (_) { chars = null; }
    if (!chars || chars.length < 1) return;

    // 先頭文字のフォント情報を保持（実用上、サイズも保持）
    var firstCA = null;
    try { firstCA = chars[0].characterAttributes; } catch (_) { firstCA = null; }
    if (!firstCA) return;

    var keepFont = null;
    var keepSize = null;
    try { keepFont = firstCA.textFont; } catch (_) { keepFont = null; }
    try { keepSize = firstCA.size; } catch (_) { keepSize = null; }

    var ca = null;
    try { ca = tr.characterAttributes; } catch (_) { ca = null; }
    if (!ca) return;

    // フォントを統一
    try { if (keepFont) ca.textFont = keepFont; } catch (_) { }
    try { if (keepSize != null) ca.size = keepSize; } catch (_) { }

    // テキストカラーを黒に統一
    try {
        var black = new GrayColor();
        black.gray = 100; // K100
        ca.fillColor = black;
    } catch (_) { }

    // それ以外を初期化（できる範囲で）
    try { ca.baselineShift = 0; } catch (_) { }
    try { ca.horizontalScale = 100; } catch (_) { }
    try { ca.verticalScale = 100; } catch (_) { }
    try { ca.rotation = 0; } catch (_) { }
    try { ca.tracking = 0; } catch (_) { }

    // カーニング/行送り
    try { ca.kerningMethod = KerningMethod.METRICS; } catch (_) { }
    try { ca.autoLeading = true; } catch (_) { }

    // 文字単位の回転/変形が残るケースがあるので、個別にも念押し
    for (var i = 0; i < chars.length; i++) {
        var c = null;
        try { c = chars[i]; } catch (_) { c = null; }
        if (!c) continue;
        var c2 = null;
        try { c2 = c.characterAttributes; } catch (_) { c2 = null; }
        if (!c2) continue;

        try { if (keepFont) c2.textFont = keepFont; } catch (_) { }
        try { if (keepSize != null) c2.size = keepSize; } catch (_) { }

        try {
            var black2 = new GrayColor();
            black2.gray = 100;
            c2.fillColor = black2;
        } catch (_) { }

        try { c2.baselineShift = 0; } catch (_) { }
        try { c2.horizontalScale = 100; } catch (_) { }
        try { c2.verticalScale = 100; } catch (_) { }
        try { c2.rotation = 0; } catch (_) { }
        try { c2.tracking = 0; } catch (_) { }
        try { c2.kerningMethod = KerningMethod.METRICS; } catch (_) { }
        try { c2.autoLeading = true; } catch (_) { }
    }
}

/* =========================================
 * High precision split / 高精度分割
 * ========================================= */

function splitTextFrameHighPrecision(textFrame) {
    if (!textFrame || textFrame.typename !== "TextFrame") return;

    // スペースを残す場合、アウトラインbounds方式ではスペース位置を保持できないためフォールバックに切替
    try {
        if (optKeepSpaces) {
            splitTextFrameFallback(textFrame);
            return;
        }
    } catch (_) { }

    var tr = null;
    try { tr = textFrame.textRange; } catch (_) { tr = null; }
    if (!tr) return;

    var chars = null;
    try { chars = tr.characters; } catch (_) { chars = null; }
    if (!chars) return;

    var n = 0;
    try { n = chars.length; } catch (_) { n = 0; }
    if (!n || n <= 0) return;

    /* アウトライン bounds を利用 / Use outline bounds */
    var outlineInfo = null;
    try {
        outlineInfo = buildOutlineCharBounds(textFrame);
    } catch (_) {
        outlineInfo = null;
    }

    if (outlineInfo && outlineInfo.ok && outlineInfo.boundsList && outlineInfo.boundsList.length > 0) {
        var boundsList = outlineInfo.boundsList;
        var layer = textFrame.layer;

        // スペースはアウトライン側に存在しないため、boundsList は「可視文字のみ」として扱う
        // chars は元テキストの全文字（スペース含む）なので、インデックスを分けて走査する
        var made = [];
        var bi = 0; // boundsList index（可視文字用）

        for (var ci = 0; ci < n; ci++) {
            var ch = null;
            try { ch = chars[ci]; } catch (_) { ch = null; }
            if (!ch) continue;

            var content = "";
            try { content = ch.contents; } catch (_) { content = ""; }
            if (content === "") continue;

            // スペース/タブ/改行などは完全に無視（生成しない / boundsも消費しない）
            // 改行/タブは無視（生成しない / boundsも消費しない）
            if (isIgnoredSpaceChar(content)) continue;

            // 「スペースを残す」ON: スペースは直前の文字に結合（新規TextFrameは作らない / boundsも消費しない）
            if (optKeepSpaces && isSpaceChar(content)) {
                if (made.length > 0) {
                    try { made[made.length - 1].contents += content; } catch (_) { }
                }
                continue;
            }

            // 可視文字に対して bounds が不足したら、この方式は成立しないのでフォールバック
            if (bi >= boundsList.length) {
                cleanupMade(made);
                safeRemoveOutlineInfo(outlineInfo);
                splitTextFrameFallback(textFrame);
                return;
            }

            // 1文字TextFrameを作成
            var nf = null;
            try {
                nf = layer.textFrames.add();
                nf.contents = content;
            } catch (_) {
                cleanupMade(made);
                safeRemoveOutlineInfo(outlineInfo);
                splitTextFrameFallback(textFrame);
                return;
            }

            // 文字属性をコピー
            try { copyCharacterAttributes(nf, ch); } catch (_) { }

            // 元TextFrameの変形（回転/拡縮など）を適用
            try { nf.matrix = textFrame.matrix; } catch (_) { }

            // まず元フレーム近傍に置く（大外れ回避）
            try { nf.left = textFrame.left; nf.top = textFrame.top; } catch (_) { }

            // 目標 bounds に一致するように nf を移動
            try {
                moveTextFrameToMatchBounds(nf, boundsList[bi]);
            } catch (_) {
                try { nf.remove(); } catch (__) { }
                cleanupMade(made);
                safeRemoveOutlineInfo(outlineInfo);
                splitTextFrameFallback(textFrame);
                return;
            }

            made.push(nf);
            bi++;
        }

        // エリア内文字に変換（UI設定 or マージ時）
        try { if (optConvertToAreaText || optMergeAreaText) made = convertTextFramesToAreaText(made); } catch (_) { }

        // エリア内文字を連結（スレッド化）
        try { if (optMergeAreaText) made = threadAreaTextFrames(made); } catch (_) { }

        // グループ化（UI設定）
        try { applyGroupingByMode(made, optGroupMode); } catch (_) { }

        // 計算用アウトラインを削除
        safeRemoveOutlineInfo(outlineInfo);

        // 元のTextFrameを削除
        try { textFrame.remove(); } catch (_) { }
        return;
    }

    /* フォールバック / Fallback */
    splitTextFrameFallback(textFrame);
}

/* =========================================
 * Outline bounds builder / アウトラインboundsの取得
 * ========================================= */

function buildOutlineCharBounds(textFrame) {
    var dup = null;
    try {
        dup = textFrame.duplicate(textFrame.parent, ElementPlacement.PLACEATBEGINNING);
    } catch (_) {
        try { dup = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING); } catch (__) { dup = null; }
    }
    if (!dup) return { ok: false };

    // アウトライン化
    var outlined = null;
    try {
        outlined = dup.createOutline();
    } catch (_) {
        try { dup.remove(); } catch (__) { }
        return { ok: false };
    }

    // createOutline 後に dup 自体が残る環境があるので消す
    try { dup.remove(); } catch (_) { }

    if (!outlined) return { ok: false };

    // outlined は GroupItem で返ることが多い
    var items = [];
    collectCharLikeItems(outlined, items);

    // 取得順は不定なので、座標から読み順に並べ替える
    try { sortItemsByTextDirection(items, textFrame); } catch (_) { }

    if (items.length === 0) {
        try { outlined.remove(); } catch (_) { }
        return { ok: false };
    }

    var boundsList = [];
    for (var i = 0; i < items.length; i++) {
        try { boundsList.push(items[i].geometricBounds); } catch (_) { }
    }

    return {
        ok: boundsList.length > 0,
        outlinedRoot: outlined,
        boundsList: boundsList
    };
}

/* Collect char-like items / 文字っぽい単位を収集 */
function collectCharLikeItems(outlinedRoot, outArr) {
    if (!outlinedRoot) return;

    var direct = [];
    try {
        var pi = outlinedRoot.pageItems;
        for (var i = 0; i < pi.length; i++) direct.push(pi[i]);
    } catch (_) { }

    if (direct.length === 1 && direct[0] && direct[0].typename === "GroupItem") {
        try {
            var inner = direct[0].pageItems;
            if (inner && inner.length > 0) {
                for (var j = 0; j < inner.length; j++) outArr.push(inner[j]);
                return;
            }
        } catch (_) { }
    }

    for (var k = 0; k < direct.length; k++) outArr.push(direct[k]);
}

/**
 * アウトライン側の items を「読み順」に並べ替える（複数行対応）。
 * - まずY（行）でクラスタリングして上→下に並べる
 * - 各行の中は「左端(L)」で左→右
 * - TextFrame.matrix は参照せず、見た目の座標だけで決める
 */
function sortItemsByTextDirection(items, textFrame) {
    if (!items || items.length <= 1) return;

    var arr = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var bnd = null;
        try { bnd = it.geometricBounds; } catch (_) { bnd = null; }
        if (!bnd || bnd.length !== 4) continue;

        var L = bnd[0], T = bnd[1], R = bnd[2], B = bnd[3];
        var cx = (L + R) / 2;
        var cy = (T + B) / 2;
        var h = Math.abs(T - B);
        arr.push({ it: it, L: L, T: T, R: R, B: B, cx: cx, cy: cy, h: h, idx: i });
    }
    if (arr.length <= 1) return;

    var th = estimateRowThreshold(arr);

    // 上→下（cy降順）
    arr.sort(function (p, q) {
        var dy = q.cy - p.cy;
        if (Math.abs(dy) > 0.001) return (dy < 0) ? -1 : 1;
        var dL = p.L - q.L;
        if (Math.abs(dL) > 0.001) return (dL < 0) ? -1 : 1;
        return p.idx - q.idx;
    });

    var rows = [];
    for (var a = 0; a < arr.length; a++) {
        var cur = arr[a];
        var placed = false;
        for (var r = 0; r < rows.length; r++) {
            if (Math.abs(cur.cy - rows[r].cy) <= th) {
                rows[r].items.push(cur);
                rows[r].cy = (rows[r].cy * (rows[r].items.length - 1) + cur.cy) / rows[r].items.length;
                placed = true;
                break;
            }
        }
        if (!placed) rows.push({ cy: cur.cy, items: [cur] });
    }

    rows.sort(function (p, q) {
        var dy2 = q.cy - p.cy;
        if (Math.abs(dy2) > 0.001) return (dy2 < 0) ? -1 : 1;
        return 0;
    });

    var out = [];
    for (var ri = 0; ri < rows.length; ri++) {
        var rowItems = rows[ri].items;
        rowItems.sort(function (p, q) {
            var dL2 = p.L - q.L;
            if (Math.abs(dL2) > 0.5) return (dL2 < 0) ? -1 : 1;

            var dT = p.T - q.T;
            if (Math.abs(dT) > 0.5) return (dT < 0) ? 1 : -1;

            var dcx = p.cx - q.cx;
            if (Math.abs(dcx) > 0.5) return (dcx < 0) ? -1 : 1;

            var dcy = p.cy - q.cy;
            if (Math.abs(dcy) > 0.5) return (dcy < 0) ? 1 : -1;

            return p.idx - q.idx;
        });

        for (var j = 0; j < rowItems.length; j++) out.push(rowItems[j].it);
    }

    for (var k = 0; k < out.length; k++) items[k] = out[k];
}

function estimateRowThreshold(arr) {
    var hs = [];
    for (var i = 0; i < arr.length; i++) {
        if (arr[i].h && arr[i].h > 0) hs.push(arr[i].h);
    }
    if (hs.length === 0) return 8;
    hs.sort(function (a, b) { return a - b; });
    var mid = hs[Math.floor(hs.length / 2)];
    if (!mid || mid <= 0) mid = hs[0];
    var th = mid * 0.6;
    if (th < 2) th = 2;
    return th;
}

function isSpaceChar(s) {
    try {
        return (s === " " || s === "\u3000"); // 半角/全角スペース
    } catch (_) {
        return false;
    }
}

function isIgnoredSpaceChar(s) {
    try {
        // 改行（段落終端）は常に無視
        if (s === "\r" || s === "\n") return true;

        // 「スペースを残す」ONのときは半角/全角スペースは無視しない
        if (optKeepSpaces && (s === " " || s === "\u3000")) return false;

        // 既定（OFF）: 半角/全角スペース/タブを無視
        return (s === " " || s === "\t" || s === "\u3000");
    } catch (_) {
        return false;
    }
}

/* Utils / ユーティリティ */

function safeRemoveOutlineInfo(info) {
    if (!info) return;
    try { if (info.outlinedRoot) info.outlinedRoot.remove(); } catch (_) { }
}

function cleanupMade(arr) {
    for (var i = 0; i < arr.length; i++) {
        try { arr[i].remove(); } catch (_) { }
    }
}

/* Match bounds by outlining / アウトライン化してbounds一致 */
function moveTextFrameToMatchBounds(nf, targetBounds) {
    if (!nf || nf.typename !== "TextFrame") return;
    if (!targetBounds || targetBounds.length !== 4) return;

    var dup = null;
    try {
        dup = nf.duplicate(nf.parent, ElementPlacement.PLACEATBEGINNING);
    } catch (_) {
        try { dup = nf.duplicate(nf.layer, ElementPlacement.PLACEATBEGINNING); } catch (__) { dup = null; }
    }
    if (!dup) return;

    var outlined = null;
    try {
        outlined = dup.createOutline();
    } catch (_) {
        try { dup.remove(); } catch (__) { }
        return;
    }

    try { dup.remove(); } catch (_) { }

    if (!outlined) {
        try { outlined.remove(); } catch (_) { }
        return;
    }

    var bNow = null;
    try { bNow = outlined.geometricBounds; } catch (_) { bNow = null; }
    try { outlined.remove(); } catch (_) { }

    if (!bNow || bNow.length !== 4) return;

    var cNow = boundsCenter(bNow);
    var cTar = boundsCenter(targetBounds);

    var dx = cTar[0] - cNow[0];
    var dy = cTar[1] - cNow[1];

    try {
        nf.left += dx;
        nf.top += dy;
    } catch (_) {
        try { nf.translate(dx, dy); } catch (__) { }
    }
}

function boundsCenter(b) {
    var cx = (b[0] + b[2]) / 2;
    var cy = (b[1] + b[3]) / 2;
    return [cx, cy];
}

/* Copy character attributes / 文字属性コピー */
function copyCharacterAttributes(dstTextFrame, srcCharacter) {
    if (!dstTextFrame || dstTextFrame.typename !== "TextFrame") return;
    if (!srcCharacter) return;

    var src = null;
    try { src = srcCharacter.characterAttributes; } catch (_) { src = null; }
    if (!src) return;

    var dst = null;
    try { dst = dstTextFrame.textRange.characterAttributes; } catch (_) { dst = null; }
    if (!dst) return;

    try { dst.textFont = src.textFont; } catch (_) { }
    try { dst.size = src.size; } catch (_) { }
    try { dst.horizontalScale = src.horizontalScale; } catch (_) { }
    try { dst.verticalScale = src.verticalScale; } catch (_) { }
    try { dst.tracking = src.tracking; } catch (_) { }
    try { dst.baselineShift = src.baselineShift; } catch (_) { }
    try { dst.rotation = src.rotation; } catch (_) { }

    try {
        if (src.fillColor && src.fillColor.typename !== "NoColor") dst.fillColor = src.fillColor;
    } catch (_) { }
    try {
        if (src.strokeColor && src.strokeColor.typename !== "NoColor") {
            dst.strokeColor = src.strokeColor;
            dst.strokeWeight = src.strokeWeight;
        }
    } catch (_) { }

    try { dst.autoLeading = src.autoLeading; } catch (_) { }
    try { if (!src.autoLeading) dst.leading = src.leading; } catch (_) { }
    try { dst.kerningMethod = src.kerningMethod; } catch (_) { }
}

/* フォールバック: 幅を積算して配置 / Fallback: accumulate widths */
function splitTextFrameFallback(textFrame) {
    if (!textFrame || textFrame.typename !== "TextFrame") return;

    var textLength = 0;
    try { textLength = textFrame.textRange.characters.length; } catch (_) { textLength = 0; }
    if (!textLength) return;

    var layer = null;
    try { layer = textFrame.layer; } catch (_) { layer = null; }
    if (!layer) return;

    // 「スペースを残す」ONのときは、左→右に走査してスペースを直前の文字に結合する
    if (optKeepSpaces) {
        var lastFrame = null;
        var made2 = [];

        for (var i2 = 0; i2 < textLength; i2++) {
            try {
                var ch2 = textFrame.textRange.characters[i2];
                var s2 = ch2.contents;

                // 改行/タブは常に無視
                if (isIgnoredSpaceChar(s2)) continue;

                // スペースは直前の文字に結合
                if (isSpaceChar(s2)) {
                    if (lastFrame) {
                        try { lastFrame.contents += s2; } catch (_) { }
                    }
                    continue;
                }

                // 通常文字: TextFrameを作成して配置
                var charAttributes2 = {
                    textFont: ch2.characterAttributes.textFont,
                    size: ch2.characterAttributes.size,
                    horizontalScale: ch2.characterAttributes.horizontalScale,
                    verticalScale: ch2.characterAttributes.verticalScale,
                    tracking: ch2.characterAttributes.tracking,
                    baselineShift: ch2.characterAttributes.baselineShift,
                    rotation: ch2.characterAttributes.rotation,
                    fillColor: (ch2.characterAttributes.fillColor.typename === "NoColor") ? null : ch2.characterAttributes.fillColor,
                    strokeColor: (ch2.characterAttributes.strokeColor.typename === "NoColor") ? null : ch2.characterAttributes.strokeColor,
                    strokeWeight: ch2.characterAttributes.strokeWeight,
                    autoLeading: ch2.characterAttributes.autoLeading,
                    leading: ch2.characterAttributes.leading,
                    kerningMethod: ch2.characterAttributes.kerningMethod
                };

                var matrix2 = textFrame.matrix;

                var newFrame2 = layer.textFrames.add();
                newFrame2.contents = s2;
                applyAttributes(newFrame2, charAttributes2);

                newFrame2.top = textFrame.top;
                newFrame2.left = textFrame.left;
                newFrame2.matrix = matrix2;

                // 先行文字（スペース含む）の幅を積算してオフセットを作る
                var offsetX2 = 0;
                for (var j2 = 0; j2 < i2; j2++) {
                    var prev2 = textFrame.textRange.characters[j2];
                    var pv2 = prev2.contents;

                    // 改行/タブは無視
                    if (pv2 === "\r" || pv2 === "\n" || pv2 === "\t") continue;

                    var tmp2 = layer.textFrames.add();
                    tmp2.contents = pv2;
                    applyAttributes(tmp2, {
                        textFont: prev2.characterAttributes.textFont,
                        size: prev2.characterAttributes.size,
                        horizontalScale: prev2.characterAttributes.horizontalScale,
                        verticalScale: prev2.characterAttributes.verticalScale,
                        tracking: prev2.characterAttributes.tracking
                    });
                    offsetX2 += tmp2.width;
                    tmp2.remove();
                }

                var angle2 = getRotationFromMatrix(matrix2);
                var rad2 = angle2 * Math.PI / 180;
                var cos2 = Math.cos(rad2);
                var sin2 = Math.sin(rad2);

                newFrame2.left = textFrame.left + offsetX2 * cos2;
                newFrame2.top = textFrame.top - offsetX2 * sin2;

                lastFrame = newFrame2;
                try { made2.push(newFrame2); } catch (_) { }

            } catch (_) { }
        }

        try { if (optConvertToAreaText || optMergeAreaText) made2 = convertTextFramesToAreaText(made2); } catch (_) { }
        try { if (optMergeAreaText) made2 = threadAreaTextFrames(made2); } catch (_) { }
        try { applyGroupingByMode(made2, optGroupMode); } catch (_) { }
        try { textFrame.remove(); } catch (_) { }
        return;
    }

    var made = [];
    for (var i = textLength - 1; i >= 0; i--) {
        try {
            var character = textFrame.textRange.characters[i];
            var charContent = character.contents;

            if (isIgnoredSpaceChar(charContent)) continue;

            var charAttributes = {
                textFont: character.characterAttributes.textFont,
                size: character.characterAttributes.size,
                horizontalScale: character.characterAttributes.horizontalScale,
                verticalScale: character.characterAttributes.verticalScale,
                tracking: character.characterAttributes.tracking,
                baselineShift: character.characterAttributes.baselineShift,
                rotation: character.characterAttributes.rotation,
                fillColor: (character.characterAttributes.fillColor.typename === "NoColor") ? null : character.characterAttributes.fillColor,
                strokeColor: (character.characterAttributes.strokeColor.typename === "NoColor") ? null : character.characterAttributes.strokeColor,
                strokeWeight: character.characterAttributes.strokeWeight,
                autoLeading: character.characterAttributes.autoLeading,
                leading: character.characterAttributes.leading,
                kerningMethod: character.characterAttributes.kerningMethod
            };

            var matrix = textFrame.matrix;

            var newFrame = layer.textFrames.add();
            newFrame.contents = charContent;

            applyAttributes(newFrame, charAttributes);

            newFrame.top = textFrame.top;
            newFrame.left = textFrame.left;
            newFrame.matrix = matrix;

            var offsetX = 0;
            for (var j = 0; j < i; j++) {
                var prevChar = textFrame.textRange.characters[j];
                if (isIgnoredSpaceChar(prevChar.contents)) continue;
                var tempFrame = layer.textFrames.add();
                tempFrame.contents = prevChar.contents;
                applyAttributes(tempFrame, {
                    textFont: prevChar.characterAttributes.textFont,
                    size: prevChar.characterAttributes.size,
                    horizontalScale: prevChar.characterAttributes.horizontalScale,
                    verticalScale: prevChar.characterAttributes.verticalScale,
                    tracking: prevChar.characterAttributes.tracking
                });
                offsetX += tempFrame.width;
                tempFrame.remove();
            }

            var angle = getRotationFromMatrix(matrix);
            var rad = angle * Math.PI / 180;
            var cosA = Math.cos(rad);
            var sinA = Math.sin(rad);

            newFrame.left = textFrame.left + offsetX * cosA;
            newFrame.top = textFrame.top - offsetX * sinA;

            try { made.push(newFrame); } catch (_) { }

        } catch (_) { }
    }

    try { if (optConvertToAreaText || optMergeAreaText) made = convertTextFramesToAreaText(made); } catch (_) { }
    try { if (optMergeAreaText) made = threadAreaTextFrames(made); } catch (_) { }
    try { applyGroupingByMode(made, optGroupMode); } catch (_) { }
    try { textFrame.remove(); } catch (_) { }
}

function applyAttributes(textFrame, attr) {
    var charAttr = null;
    try { charAttr = textFrame.textRange.characterAttributes; } catch (_) { charAttr = null; }
    if (!charAttr) return;

    try {
        if (attr.textFont) charAttr.textFont = attr.textFont;
        if (attr.size) charAttr.size = attr.size;
        charAttr.horizontalScale = (attr.horizontalScale !== undefined) ? attr.horizontalScale : 100;
        charAttr.verticalScale = (attr.verticalScale !== undefined) ? attr.verticalScale : 100;
        charAttr.tracking = (attr.tracking !== undefined) ? attr.tracking : 0;
        charAttr.baselineShift = (attr.baselineShift !== undefined) ? attr.baselineShift : 0;
        charAttr.rotation = (attr.rotation !== undefined) ? attr.rotation : 0;

        if (attr.fillColor) charAttr.fillColor = attr.fillColor;
        if (attr.strokeColor) {
            charAttr.strokeColor = attr.strokeColor;
            charAttr.strokeWeight = attr.strokeWeight;
        }

        if (attr.autoLeading !== undefined) {
            charAttr.autoLeading = attr.autoLeading;
            if (!attr.autoLeading && attr.leading) {
                charAttr.leading = attr.leading;
            }
        }

        if (attr.kerningMethod) charAttr.kerningMethod = attr.kerningMethod;
    } catch (_) { }
}

function getRotationFromMatrix(matrix) {
    try {
        return Math.atan2(matrix.mValueB, matrix.mValueA) * 180 / Math.PI;
    } catch (_) {
        return 0;
    }
}

/* =========================================
 * Area text conversion / エリア内文字への変換
 * ========================================= */

function convertTextFramesToAreaText(frames) {
    if (!frames || frames.length === 0) return frames;

    // 変換結果のTextFrame配列を返す（失敗時は元を返す）
    var out = [];

    // selection を汚さないよう退避
    var doc = null;
    try { doc = app.activeDocument; } catch (_) { doc = null; }
    var oldSel = null;
    try { oldSel = doc ? doc.selection : null; } catch (_) { oldSel = null; }

    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        if (!tf || tf.typename !== "TextFrame") {
            continue;
        }

        // すでにエリア内文字ならそのまま
        try {
            if (tf.kind === TextType.AREATEXT) {
                out.push(tf);
                continue;
            }
        } catch (_) { }

        // ポイントテキストなら、ネイティブAPIでエリア内文字へ変換（最優先）
        try {
            if (tf.kind === TextType.POINTTEXT && tf.convertPointObjectToAreaObject) {
                tf.convertPointObjectToAreaObject();
            }
        } catch (_) { }

        // 変換できたらそのまま採用
        try {
            if (tf.kind === TextType.AREATEXT) {
                out.push(tf);
                continue;
            }
        } catch (_) { }

        // まずはメニューコマンドで変換を試す（最も見た目を保持できる）
        var converted = null;
        try {
            if (doc) {
                doc.selection = [tf];
                try { app.executeMenuCommand('ConvertToAreaType'); } catch (_) {
                    // 環境差のため別名も試す
                    try { app.executeMenuCommand('ConvertToAreaText'); } catch (__) { }
                }
                try {
                    if (doc.selection && doc.selection.length === 1 && doc.selection[0].typename === "TextFrame") {
                        converted = doc.selection[0];
                    }
                } catch (_) { }
            }
        } catch (_) {
            converted = null;
        }

        if (converted && converted.typename === "TextFrame") {
            out.push(converted);
            continue;
        }

        // うまくいかなければ簡易フォールバック: bounds を枠にして areaText を作る
        try {
            var b = tf.geometricBounds; // [L,T,R,B]
            var L = b[0], T = b[1], R = b[2], B = b[3];
            var w = R - L;
            var h = T - B;
            if (w <= 0 || h <= 0) {
                out.push(tf);
                continue;
            }

            var layer = tf.layer;
            var rect = layer.pathItems.rectangle(T, L, w, h);
            rect.stroked = false;
            rect.filled = false;

            var at = null;
            try { at = layer.textFrames.areaText(rect); } catch (_) { at = null; }
            if (!at) {
                try { rect.remove(); } catch (_) { }
                out.push(tf);
                continue;
            }

            // 内容と主要属性を移植
            try { at.contents = tf.contents; } catch (_) { }
            try { at.matrix = tf.matrix; } catch (_) { }

            try {
                var srcCA = tf.textRange.characterAttributes;
                var dstCA = at.textRange.characterAttributes;
                // フォント/サイズ/スケール/トラッキング等（可能な範囲）
                try { dstCA.textFont = srcCA.textFont; } catch (_) { }
                try { dstCA.size = srcCA.size; } catch (_) { }
                try { dstCA.horizontalScale = srcCA.horizontalScale; } catch (_) { }
                try { dstCA.verticalScale = srcCA.verticalScale; } catch (_) { }
                try { dstCA.tracking = srcCA.tracking; } catch (_) { }
                try { dstCA.baselineShift = srcCA.baselineShift; } catch (_) { }
                try { dstCA.rotation = srcCA.rotation; } catch (_) { }
                try { dstCA.kerningMethod = srcCA.kerningMethod; } catch (_) { }
                try { dstCA.autoLeading = srcCA.autoLeading; } catch (_) { }
                try { if (!srcCA.autoLeading) dstCA.leading = srcCA.leading; } catch (_) { }
                try { if (srcCA.fillColor && srcCA.fillColor.typename !== "NoColor") dstCA.fillColor = srcCA.fillColor; } catch (_) { }
                try { if (srcCA.strokeColor && srcCA.strokeColor.typename !== "NoColor") { dstCA.strokeColor = srcCA.strokeColor; dstCA.strokeWeight = srcCA.strokeWeight; } } catch (_) { }
            } catch (_) { }

            // 元を削除（rect は areaText の枠として保持される）
            try { tf.remove(); } catch (_) { }

            out.push(at);
            continue;

        } catch (_) {
            // 失敗したらそのまま
            out.push(tf);
        }
    }

    // selection 復元
    try { if (doc) doc.selection = oldSel; } catch (_) { }

    return out;
}

/* =========================================
 * Area text threading / エリア内文字のスレッド化
 * ========================================= */

function threadAreaTextFrames(frames) {
    if (!frames || frames.length < 2) return frames;

    var doc = null;
    try { doc = app.activeDocument; } catch (_) { doc = null; }
    if (!doc) return frames;

    // 読み順で並べ替え（複数行対応）
    var rows = [];
    try { rows = framesToRowsInReadingOrder(frames); } catch (_) { rows = []; }

    var ordered = [];
    if (rows && rows.length > 0) {
        for (var r = 0; r < rows.length; r++) {
            for (var i = 0; i < rows[r].length; i++) ordered.push(rows[r][i]);
        }
    } else {
        // フォールバック: 入力順
        for (var j = 0; j < frames.length; j++) ordered.push(frames[j]);
    }

    // エリア内文字だけに絞る（念のため）
    var area = [];
    for (var k = 0; k < ordered.length; k++) {
        var tf = ordered[k];
        if (!tf || tf.typename !== "TextFrame") continue;
        try {
            if (tf.kind === TextType.AREATEXT) area.push(tf);
        } catch (_) { }
    }
    if (area.length < 2) return frames;

    // selection を退避して threadTextCreate を実行
    var oldSel = null;
    try { oldSel = doc.selection; } catch (_) { oldSel = null; }

    try {
        doc.selection = area;
        // スレッド（連結）
        app.executeMenuCommand('threadTextCreate');
    } catch (_) {
        // 失敗しても無視
    }

    try { doc.selection = oldSel; } catch (_) { }

    // 連結後も参照はそのまま使えるため、並び順（area）を返す
    return area;
}

/* =========================================
 * Grouping helpers / グループ化ヘルパー
 * ========================================= */

function applyGroupingByMode(frames, mode) {
    if (!frames || frames.length < 2) return;
    if (!mode || mode === "none") return;

    // 読み順で並べ替えた行リストを作る
    var rows = framesToRowsInReadingOrder(frames);
    if (rows.length === 0) return;

    if (mode === "all") {
        // 全体を1グループ
        var flat = [];
        for (var r = 0; r < rows.length; r++) {
            for (var i = 0; i < rows[r].length; i++) flat.push(rows[r][i]);
        }
        groupItems(flat);
        return;
    }

    if (mode === "line") {
        // 各行ごとにグループ
        for (var rr = 0; rr < rows.length; rr++) {
            if (rows[rr].length >= 2) groupItems(rows[rr]);
        }
        return;
    }
}

function groupItems(items) {
    if (!items || items.length < 2) return null;

    var parent = null;
    try { parent = items[0].parent; } catch (_) { parent = null; }
    if (!parent) return null;

    var g = null;
    try {
        // 同階層にグループを作る
        if (parent.typename === "Layer") g = parent.groupItems.add();
        else if (parent.groupItems) g = parent.groupItems.add();
        else g = items[0].layer.groupItems.add();
    } catch (_) {
        try { g = items[0].layer.groupItems.add(); } catch (__) { g = null; }
    }
    if (!g) return null;

    // 読み順のままグループへ移動
    for (var i = 0; i < items.length; i++) {
        try { items[i].move(g, ElementPlacement.PLACEATEND); } catch (_) { }
    }

    return g;
}

function framesToRowsInReadingOrder(frames) {
    var arr = [];
    for (var i = 0; i < frames.length; i++) {
        var it = frames[i];
        if (!it) continue;
        var b = null;
        try { b = it.geometricBounds; } catch (_) { b = null; }
        if (!b || b.length !== 4) continue;

        var L = b[0], T = b[1], R = b[2], B = b[3];
        var cx = (L + R) / 2;
        var cy = (T + B) / 2;
        var h = Math.abs(T - B);
        arr.push({ it: it, L: L, T: T, R: R, B: B, cx: cx, cy: cy, h: h, idx: i });
    }
    if (arr.length === 0) return [];

    var th = estimateRowThreshold(arr);

    // 上→下（cy降順）に並べて行分け
    arr.sort(function (p, q) {
        var dy = q.cy - p.cy;
        if (Math.abs(dy) > 0.001) return (dy < 0) ? -1 : 1;
        var dL = p.L - q.L;
        if (Math.abs(dL) > 0.001) return (dL < 0) ? -1 : 1;
        return p.idx - q.idx;
    });

    var rows = [];
    for (var a = 0; a < arr.length; a++) {
        var cur = arr[a];
        var placed = false;
        for (var r = 0; r < rows.length; r++) {
            if (Math.abs(cur.cy - rows[r].cy) <= th) {
                rows[r].items.push(cur);
                rows[r].cy = (rows[r].cy * (rows[r].items.length - 1) + cur.cy) / rows[r].items.length;
                placed = true;
                break;
            }
        }
        if (!placed) rows.push({ cy: cur.cy, items: [cur] });
    }

    // 行を上→下
    rows.sort(function (p, q) {
        var dy2 = q.cy - p.cy;
        if (Math.abs(dy2) > 0.001) return (dy2 < 0) ? -1 : 1;
        return 0;
    });

    // 行内 左→右
    var outRows = [];
    for (var ri = 0; ri < rows.length; ri++) {
        var rowItems = rows[ri].items;
        rowItems.sort(function (p, q) {
            var dL2 = p.L - q.L;
            if (Math.abs(dL2) > 0.5) return (dL2 < 0) ? -1 : 1;

            var dT = p.T - q.T;
            if (Math.abs(dT) > 0.5) return (dT < 0) ? 1 : -1;

            return p.idx - q.idx;
        });

        var row = [];
        for (var j = 0; j < rowItems.length; j++) row.push(rowItems[j].it);
        outRows.push(row);
    }

    return outRows;
}