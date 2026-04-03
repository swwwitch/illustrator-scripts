
#target illustrator
#targetengine "SelectionInspector"

// Keep a global reference so the palette does not get garbage-collected.
// Re-running the script will reuse the same palette.
var __SelectionInspector_Palette__ = this.__SelectionInspector_Palette__;
var __SI_lastDocErrAt__ = this.__SI_lastDocErrAt__ || 0;
this.__SI_lastDocErrAt__ = __SI_lastDocErrAt__;
// --- Idle coordination globals
var __SI_idleTask__ = this.__SI_idleTask__;
var __SI_needRefresh__ = this.__SI_needRefresh__ || false;
var __SI_needExport__ = this.__SI_needExport__ || false;
this.__SI_needRefresh__ = __SI_needRefresh__;
this.__SI_needExport__ = __SI_needExport__;


/*

### スクリプト名：

SelectionInspector.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionInspector.jsx

### 概要：

- 選択中／全体のオブジェクト数をカウントし、2カラムのパレット（modeless）で表示
- テキスト・配置画像・透明・グループ・パス・ガイドの統計を表示
- 「書き出し」でレポート（テキスト）を書き出し可能

### 主な機能：

- オブジェクト数（選択／全体）・アートボード数を表示
- テキスト（ポイント／エリア／パス上）・文字数・段落数・強制改行数を表示
- 配置画像（リンク／埋め込み／リンク切れ）を表示
- 透明（不透明度<100、描画モード≠通常）を表示
- グループ／クリップグループを表示
- パス（オープン／クローズ／アンカー／ハンドル／複合パス／複合シェイプ）を表示
  ※ ガイド（guides=true）はパス統計から除外
- ガイド（ルーラー／アートボード／その他）を表示
- 「書き出し」ボタンで集計結果をテキストファイルとして保存

### 処理の流れ：

- ドキュメントと選択オブジェクトを解析し、各種情報を取得
- パレットに2カラムで情報を整然と表示
- 書き出し機能でデスクトップにテキスト保存

### 更新履歴：

- v1.0 (20250806) : 初期バージョン
- v1.1 (20250806) : 書き出し機能を追加
- v1.2 (20250807) : UI調整
- v1.3 (20260301) : パスのハンドル数を追加
- v1.4 (20260301) : 強制改行（ソフトリターン）数を追加
- v1.4.1 (20260302) : UI調整（OK→閉じる、キャンセル削除、書き出しを左へ移動）
- v1.5 (20260302) : ガイド集計（ルーラー/アートボード/その他）、ガイド除外のパス統計、パネル配置と書き出しレイアウト更新
- v1.6 (20260304) : パレット（閉じずに使える）に変更

---

### Script Name:

SelectionInspector.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/SelectionInspector.jsx

### Overview:

- Count selection / document totals and show them in a two-column palette (modeless)
- Show stats for text, images, transparency, groups, paths, and guides
- Export the stats as a plain text report

### Main Features:

- Show object counts (Selection / All) and artboard count
- Text stats: point/area/path text, characters, paragraphs, forced line breaks
- Image stats: linked/embedded/broken links
- Transparency stats: opacity < 100, blend mode ≠ Normal
- Group stats: groups and clipping groups
- Path stats: open/closed, anchor points, handles, compound paths, compound shapes
  * Guide paths (guides=true) are excluded from path stats
- Guide breakdown: ruler / artboard / other
- Export button writes a plain text report to the desktop

### Process Flow:

- Analyze document and selected objects
- Display data in a two-column palette UI
- Export stats to desktop text file

### Changelog:

- v1.0 (20250806): Initial version
- v1.1 (20250806): Added export feature
- v1.2 (20250807): UI adjustments
- v1.3 (20260301): Added handle count for paths
- v1.4 (20260301): Added forced line break (soft return) count
- v1.4.1 (20260302): UI tweaks (OK→Close, removed Cancel, moved Export to left)
- v1.5 (20260302): Added guide breakdown (ruler/artboard/other), excluded guides from path stats, updated panel arrangement and export layout
- v1.6 (20260304): Switched to a modeless palette (usable without closing)

*/

var SCRIPT_VERSION = "v1.6";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "選択／全体オブジェクトのカウント " + SCRIPT_VERSION,
        en: "Selection / All Objects Count " + SCRIPT_VERSION
    },
    other: {
        ja: "基本",
        en: "Basics"
    },
    artboards: {
        ja: "アートボード：",
        en: "Artboards:"
    },
    objects: {
        ja: "オブジェクト：",
        en: "Objects:"
    },
    charPara: {
        ja: "文字・段落",
        en: "Characters & Paragraphs"
    },
    transparency: {
        ja: "透明",
        en: "Transparency"
    },
    chars: {
        ja: "文字数：",
        en: "Characters:"
    },
    paras: {
        ja: "段落数：",
        en: "Paragraphs:"
    },
    forcedBreaks: {
        ja: "強制改行：",
        en: "Line Breaks:"
    },
    opacityLt100: {
        ja: "不透明度<100：",
        en: "Opacity < 100:"
    },
    blendNotNormal: {
        ja: "描画モード≠通常：",
        en: "Blend Mode ≠ Normal:"
    },
    texts: {
        ja: "テキスト：",
        en: "Text Frames:"
    },
    pointText: {
        ja: "ポイント文字：",
        en: "Point Type:"
    },
    areaText: {
        ja: "エリア内文字：",
        en: "Area Type:"
    },
    pathText: {
        ja: "パス上文字：",
        en: "Path Text:"
    },
    link: {
        ja: "配置画像",
        en: "Images"
    },
    linked: {
        ja: "リンク：",
        en: "Linked Images:"
    },
    embed: {
        ja: "埋め込み：",
        en: "Embedded Images:"
    },
    broken: {
        ja: "リンク切れ：",
        en: "Broken Links:"
    },
    group: {
        ja: "グループ：",
        en: "Groups:"
    },
    clipGroup: {
        ja: "クリップグループ：",
        en: "Clipping Groups:"
    },
    path: {
        ja: "パス",
        en: "Paths"
    },
    guide: {
        ja: "ガイド",
        en: "Guides"
    },
    rulerGuides: {
        ja: "ルーラーガイド：",
        en: "Ruler Guides:"
    },
    artboardGuides: {
        ja: "アートボードガイド：",
        en: "Artboard Guides:"
    },
    otherGuides: {
        ja: "その他のガイド：",
        en: "Other Guides:"
    },
    pathCount: {
        ja: "パス：",
        en: "Paths:"
    },
    openPath: {
        ja: "オープンパス：",
        en: "Open Paths:"
    },
    closedPath: {
        ja: "クローズパス：",
        en: "Closed Paths:"
    },
    anchors: {
        ja: "アンカーポイント：",
        en: "Anchor Points:"
    },
    handles: {
        ja: "ハンドル：",
        en: "Handles:"
    },
    compoundPath: {
        ja: "複合パス：",
        en: "Compound Paths:"
    },
    compoundShape: {
        ja: "複合シェイプ：",
        en: "Compound Shapes:"
    },
    ok: {
        ja: "閉じる",
        en: "Close"
    },
    exportPreset: {
        ja: "書き出し",
        en: "Export"
    },
    refresh: {
        ja: "更新",
        en: "Refresh"
    }
};

function main() {
    try {
        function getActiveDocSafe() {
            // Try to get a live document handle in the most permissive way.
            // Some Illustrator states throw on activeDocument even when documents exist.
            var n = 0;
            try { n = app.documents.length; } catch (eLen) { n = 0; }
            if (n <= 0) return null;

            // 1) Prefer activeDocument
            try {
                return app.activeDocument;
            } catch (e0) { }

            // 2) Try to force-activate the first document
            try {
                app.activeDocument = app.documents[0];
            } catch (e1) { }
            try {
                return app.activeDocument;
            } catch (e2) { }

            // 3) Fallback: return the first document reference
            try {
                return app.documents[0];
            } catch (e3) {
                return null;
            }
        }

        var doc = getActiveDocSafe();
        if (!doc) {
            alert(lang === "ja" ? "ドキュメントが開かれていません。" : "No document is open.");
            return;
        }

        /* 選択オブジェクトを取得 / Get selected objects */
        var sel = [];
        try {
            sel = doc.selection;
            if (!sel) sel = [];
        } catch (eSel) {
            sel = [];
        }

        function safeLen(x) {
            try { return x && x.length ? x.length : 0; } catch (e) { return 0; }
        }



        function computeStats(doc, sel) {
            // Normalize selection
            if (!sel) sel = [];

            // Selected object count
            var totalObjSel = sel.length;

            // Document total object count (recursive)
            var totalObjAll = 0;
            function countAllObjects(items) {
                for (var i = 0; i < items.length; i++) {
                    var it = items[i];
                    totalObjAll++;
                    if (it.typename === "GroupItem") {
                        countAllObjects(it.pageItems);
                    } else if (it.typename === "CompoundPathItem") {
                        countAllObjects(it.pathItems);
                    }
                }
            }
            countAllObjects(doc.pageItems);

            // Basic
            var artboardCount = 0;
            try { artboardCount = doc.artboards.length; } catch (e) { artboardCount = 0; }

            // Type counters
            var textCountSel = 0, textCountAll = 0;
            var pointTextSel = 0, areaTextSel = 0, pathTextSel = 0;
            var pointTextAll = 0, areaTextAll = 0, pathTextAll = 0;

            var totalCharSel = 0, totalCharAll = 0;
            var paraCountSel = 0, paraCountAll = 0;
            var forcedBreakSel = 0, forcedBreakAll = 0;

            var opacityLt100Sel = 0, opacityLt100All = 0;
            var blendNotNormalSel = 0, blendNotNormalAll = 0;

            var linkedSel = 0, linkedAll = 0;
            var embeddedSel = 0, embeddedAll = 0;
            var brokenLinkSel = 0, brokenLinkAll = 0;

            var groupSel = 0, groupAll = 0;
            var clipGroupSel = 0, clipGroupAll = 0;

            var pathCountSel = 0, pathCountAll = 0;
            var openPathSel = 0, openPathAll = 0;
            var closedPathSel = 0, closedPathAll = 0;
            var anchorCountSel = 0, anchorCountAll = 0;
            var handleCountSel = 0, handleCountAll = 0;
            var compoundPathSel = 0, compoundPathAll = 0;
            var compoundShapeSel = 0, compoundShapeAll = 0;

            var rulerGuideSel = 0, rulerGuideAll = 0;
            var artboardGuideSel = 0, artboardGuideAll = 0;
            var otherGuideSel = 0, otherGuideAll = 0;

            function __getMaxArtboardSpan(doc) {
                var maxSpan = 0;
                try {
                    var abs = Math.abs;
                    var ab = doc.artboards;
                    for (var i = 0; i < ab.length; i++) {
                        var r = ab[i].artboardRect;
                        var w = abs(r[2] - r[0]);
                        var h = abs(r[1] - r[3]);
                        if (w > maxSpan) maxSpan = w;
                        if (h > maxSpan) maxSpan = h;
                    }
                } catch (e) { }
                return maxSpan;
            }

            function __isArtboardGuidePathItem(pi, doc) {
                try {
                    if (!pi || pi.typename !== "PathItem") return false;
                    if (!pi.guides) return false;
                    if (pi.closed) return false;
                    if (!pi.pathPoints || pi.pathPoints.length !== 2) return false;

                    var a0 = pi.pathPoints[0].anchor;
                    var a1 = pi.pathPoints[1].anchor;
                    var T = 0.01;
                    var sameX = Math.abs(a0[0] - a1[0]) <= T;
                    var sameY = Math.abs(a0[1] - a1[1]) <= T;
                    if (!(sameX || sameY)) return false;

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
                        if (abs(len - w) <= tol) return true;
                        if (abs(len - h) <= tol) return true;
                    }
                    return false;
                } catch (e) {
                    return false;
                }
            }

            function __isRulerGuidePathItem(pi) {
                try {
                    if (!pi || pi.typename !== "PathItem") return false;
                    if (!pi.guides) return false;
                    if (pi.closed) return false;
                    if (!pi.pathPoints || pi.pathPoints.length !== 2) return false;

                    var a0 = pi.pathPoints[0].anchor;
                    var a1 = pi.pathPoints[1].anchor;
                    var T = 0.01;

                    var sameX = Math.abs(a0[0] - a1[0]) <= T;
                    var sameY = Math.abs(a0[1] - a1[1]) <= T;
                    if (!(sameX || sameY)) return false;

                    var dx = a0[0] - a1[0];
                    var dy = a0[1] - a1[1];
                    var len = Math.sqrt(dx * dx + dy * dy);

                    var maxSpan = __getMaxArtboardSpan(doc);
                    if (!(maxSpan > 0)) return true;

                    return (len > maxSpan);
                } catch (e) {
                    return false;
                }
            }

            function __countGuidesForPathItem(pi) {
                var r = { ruler: 0, artboard: 0, other: 0 };
                try {
                    if (!pi || pi.typename !== "PathItem") return r;
                    if (!pi.guides) return r;

                    if (__isRulerGuidePathItem(pi)) {
                        r.ruler = 1;
                    } else if (__isArtboardGuidePathItem(pi, doc)) {
                        r.artboard = 1;
                    } else {
                        r.other = 1;
                    }
                } catch (e) { }
                return r;
            }

            function countTransparencyForItem(it) {
                var r = { opacityLt100: 0, blendNotNormal: 0 };
                if (!it) return r;
                try {
                    if (typeof it.opacity === "number" && it.opacity < 100) r.opacityLt100 = 1;
                } catch (e) { }
                try {
                    if (it.blendingMode !== undefined && it.blendingMode !== BlendModes.NORMAL) r.blendNotNormal = 1;
                } catch (e) { }
                return r;
            }

            function __isGuidePathItem(pi) {
                try {
                    return (pi && pi.typename === "PathItem" && pi.guides === true);
                } catch (e) {
                    return false;
                }
            }

            function countHandlesInPathItem(pi) {
                var c = 0;
                try {
                    var pts = pi.pathPoints;
                    for (var i = 0; i < pts.length; i++) {
                        var p = pts[i];
                        var a = p.anchor;
                        var l = p.leftDirection;
                        var r = p.rightDirection;
                        if (l[0] !== a[0] || l[1] !== a[1]) c++;
                        if (r[0] !== a[0] || r[1] !== a[1]) c++;
                    }
                } catch (e) { }
                return c;
            }

            function countForcedBreaksFromContents(s) {
                if (!s || !s.length) return 0;
                var m = s.match(/[\u0003\n\u2028]/g);
                return m ? m.length : 0;
            }

            // Selected loop
            for (var i = 0; i < sel.length; i++) {
                var sIt = sel[i];

                if (sIt.typename === "TextFrame") {
                    textCountSel++;
                    try { totalCharSel += safeLen(sIt.characters); } catch (e) { }
                    try { paraCountSel += safeLen(sIt.paragraphs); } catch (e) { }
                    try { forcedBreakSel += countForcedBreaksFromContents(sIt.contents); } catch (e) { }
                    try {
                        if (sIt.kind === TextType.POINTTEXT) pointTextSel++;
                        else if (sIt.kind === TextType.AREATEXT) areaTextSel++;
                        else if (sIt.kind === TextType.PATHTEXT) pathTextSel++;
                    } catch (e) { }
                }

                if (sIt.typename === "PlacedItem") {
                    if (sIt.embedded) {
                        embeddedSel++;
                    } else {
                        linkedSel++;
                        try {
                            var f = sIt.file;
                            if (!f || !f.exists) brokenLinkSel++;
                        } catch (e) {
                            brokenLinkSel++;
                        }
                    }
                }

                if (sIt.typename === "GroupItem") {
                    groupSel++;
                    try { if (sIt.clipped) clipGroupSel++; } catch (e) { }
                }

                // compound counters (selected)
                if (sIt.typename === "CompoundPathItem") {
                    compoundPathSel++;
                }
                if (sIt.typename === "PluginItem") {
                    try {
                        if (sIt.name && sIt.name.indexOf("Compound Shape") !== -1) {
                            compoundShapeSel++;
                        }
                    } catch (e) { }
                }

                // transparency (selected)
                try {
                    var t0 = countTransparencyForItem(sIt);
                    opacityLt100Sel += t0.opacityLt100;
                    blendNotNormalSel += t0.blendNotNormal;
                } catch (e) { }

                // guides (selected)
                try {
                    if (sIt.typename === "PathItem") {
                        var g0 = __countGuidesForPathItem(sIt);
                        rulerGuideSel += g0.ruler;
                        artboardGuideSel += g0.artboard;
                        otherGuideSel += g0.other;
                    } else if (sIt.typename === "CompoundPathItem") {
                        for (var gg = 0; gg < sIt.pathItems.length; gg++) {
                            var g1 = __countGuidesForPathItem(sIt.pathItems[gg]);
                            rulerGuideSel += g1.ruler;
                            artboardGuideSel += g1.artboard;
                            otherGuideSel += g1.other;
                        }
                    }
                } catch (e) { }

                // paths (selected)
                if (sIt.typename === "PathItem") {
                    if (!__isGuidePathItem(sIt)) {
                        pathCountSel++;
                        try {
                            if (sIt.closed) closedPathSel++;
                            else openPathSel++;
                        } catch (e) { }
                        try { anchorCountSel += sIt.pathPoints.length; } catch (e) { }
                        handleCountSel += countHandlesInPathItem(sIt);
                    }
                } else if (sIt.typename === "CompoundPathItem") {
                    for (var j = 0; j < sIt.pathItems.length; j++) {
                        if (__isGuidePathItem(sIt.pathItems[j])) continue;
                        pathCountSel++;
                        try {
                            if (sIt.pathItems[j].closed) closedPathSel++;
                            else openPathSel++;
                        } catch (e) { }
                        try { anchorCountSel += sIt.pathItems[j].pathPoints.length; } catch (e) { }
                        handleCountSel += countHandlesInPathItem(sIt.pathItems[j]);
                    }
                }
            }

            // All items loop
            var allItems = doc.pageItems;
            for (var k = 0; k < allItems.length; k++) {
                var obj = allItems[k];

                if (obj.typename === "TextFrame") {
                    textCountAll++;
                    try { totalCharAll += safeLen(obj.characters); } catch (e) { }
                    try { paraCountAll += safeLen(obj.paragraphs); } catch (e) { }
                    try { forcedBreakAll += countForcedBreaksFromContents(obj.contents); } catch (e) { }
                    try {
                        if (obj.kind === TextType.POINTTEXT) pointTextAll++;
                        else if (obj.kind === TextType.AREATEXT) areaTextAll++;
                        else if (obj.kind === TextType.PATHTEXT) pathTextAll++;
                    } catch (e) { }
                }

                if (obj.typename === "PlacedItem") {
                    if (obj.embedded) {
                        embeddedAll++;
                    } else {
                        linkedAll++;
                        try {
                            var fAll = obj.file;
                            if (!fAll || !fAll.exists) brokenLinkAll++;
                        } catch (e) {
                            brokenLinkAll++;
                        }
                    }
                }

                if (obj.typename === "GroupItem") {
                    groupAll++;
                    try { if (obj.clipped) clipGroupAll++; } catch (e) { }
                }

                if (obj.typename === "CompoundPathItem") {
                    compoundPathAll++;
                }
                if (obj.typename === "PluginItem") {
                    try {
                        if (obj.name && obj.name.indexOf("Compound Shape") !== -1) {
                            compoundShapeAll++;
                        }
                    } catch (e) { }
                }

                // transparency (all)
                try {
                    var tAll = countTransparencyForItem(obj);
                    opacityLt100All += tAll.opacityLt100;
                    blendNotNormalAll += tAll.blendNotNormal;
                } catch (e) { }

                // guides (all)
                try {
                    if (obj.typename === "PathItem") {
                        var gAll0 = __countGuidesForPathItem(obj);
                        rulerGuideAll += gAll0.ruler;
                        artboardGuideAll += gAll0.artboard;
                        otherGuideAll += gAll0.other;
                    } else if (obj.typename === "CompoundPathItem") {
                        for (var gg2 = 0; gg2 < obj.pathItems.length; gg2++) {
                            var gAll1 = __countGuidesForPathItem(obj.pathItems[gg2]);
                            rulerGuideAll += gAll1.ruler;
                            artboardGuideAll += gAll1.artboard;
                            otherGuideAll += gAll1.other;
                        }
                    }
                } catch (e) { }

                // paths (all)
                if (obj.typename === "PathItem") {
                    if (!__isGuidePathItem(obj)) {
                        pathCountAll++;
                        try {
                            if (obj.closed) closedPathAll++;
                            else openPathAll++;
                        } catch (e) { }
                        try { anchorCountAll += obj.pathPoints.length; } catch (e) { }
                        handleCountAll += countHandlesInPathItem(obj);
                    }
                } else if (obj.typename === "CompoundPathItem") {
                    for (var m = 0; m < obj.pathItems.length; m++) {
                        if (__isGuidePathItem(obj.pathItems[m])) continue;
                        pathCountAll++;
                        try {
                            if (obj.pathItems[m].closed) closedPathAll++;
                            else openPathAll++;
                        } catch (e) { }
                        try { anchorCountAll += obj.pathItems[m].pathPoints.length; } catch (e) { }
                        handleCountAll += countHandlesInPathItem(obj.pathItems[m]);
                    }
                }
            }

            return {
                totalObjSel: totalObjSel,
                totalObjAll: totalObjAll,
                artboardCount: artboardCount,

                textCountSel: textCountSel,
                textCountAll: textCountAll,
                pointTextSel: pointTextSel,
                pointTextAll: pointTextAll,
                areaTextSel: areaTextSel,
                areaTextAll: areaTextAll,
                pathTextSel: pathTextSel,
                pathTextAll: pathTextAll,

                totalCharSel: totalCharSel,
                totalCharAll: totalCharAll,
                paraCountSel: paraCountSel,
                paraCountAll: paraCountAll,
                forcedBreakSel: forcedBreakSel,
                forcedBreakAll: forcedBreakAll,

                opacityLt100Sel: opacityLt100Sel,
                opacityLt100All: opacityLt100All,
                blendNotNormalSel: blendNotNormalSel,
                blendNotNormalAll: blendNotNormalAll,

                linkedSel: linkedSel,
                linkedAll: linkedAll,
                embeddedSel: embeddedSel,
                embeddedAll: embeddedAll,
                brokenLinkSel: brokenLinkSel,
                brokenLinkAll: brokenLinkAll,

                groupSel: groupSel,
                groupAll: groupAll,
                clipGroupSel: clipGroupSel,
                clipGroupAll: clipGroupAll,

                pathCountSel: pathCountSel,
                pathCountAll: pathCountAll,
                openPathSel: openPathSel,
                openPathAll: openPathAll,
                closedPathSel: closedPathSel,
                closedPathAll: closedPathAll,
                anchorCountSel: anchorCountSel,
                anchorCountAll: anchorCountAll,
                handleCountSel: handleCountSel,
                handleCountAll: handleCountAll,
                compoundPathSel: compoundPathSel,
                compoundPathAll: compoundPathAll,
                compoundShapeSel: compoundShapeSel,
                compoundShapeAll: compoundShapeAll,

                rulerGuideSel: rulerGuideSel,
                rulerGuideAll: rulerGuideAll,
                artboardGuideSel: artboardGuideSel,
                artboardGuideAll: artboardGuideAll,
                otherGuideSel: otherGuideSel,
                otherGuideAll: otherGuideAll
            };
        }

        // Compute initial stats
        var currentStats = computeStats(doc, sel);

        /* 結果を表示（ダイアログボックス） / Display results (dialog box) */
        // Reuse existing palette if it already exists
        var dlg = __SelectionInspector_Palette__;
        if (dlg && dlg instanceof Window) {
            try {
                dlg.close();
            } catch (e0) { }
            dlg = null;
        }
        dlg = new Window("palette", LABELS.dialogTitle[lang]);
        __SelectionInspector_Palette__ = dlg;
        this.__SelectionInspector_Palette__ = dlg;
        dlg.orientation = "column";
        dlg.alignChildren = "center";

        var UI = {};

        function fmtPair(a, b) {
            return a + " / " + b;
        }

        function updateUIFromStats(st) {
            // Basics
            if (UI.artboards) UI.artboards.text = String(st.artboardCount);
            if (UI.objects) UI.objects.text = fmtPair(st.totalObjSel, st.totalObjAll);

            // Characters & paragraphs
            if (UI.chars) UI.chars.text = fmtPair(st.totalCharSel, st.totalCharAll);
            if (UI.paras) UI.paras.text = fmtPair(st.paraCountSel, st.paraCountAll);
            if (UI.forcedBreaks) UI.forcedBreaks.text = fmtPair(st.forcedBreakSel, st.forcedBreakAll);

            // Text frames
            if (UI.texts) UI.texts.text = fmtPair(st.textCountSel, st.textCountAll);
            if (UI.pointText) UI.pointText.text = fmtPair(st.pointTextSel, st.pointTextAll);
            if (UI.areaText) UI.areaText.text = fmtPair(st.areaTextSel, st.areaTextAll);
            if (UI.pathText) UI.pathText.text = fmtPair(st.pathTextSel, st.pathTextAll);

            // Transparency
            if (UI.opacityLt100) UI.opacityLt100.text = fmtPair(st.opacityLt100Sel, st.opacityLt100All);
            if (UI.blendNotNormal) UI.blendNotNormal.text = fmtPair(st.blendNotNormalSel, st.blendNotNormalAll);

            // Images
            if (UI.linked) UI.linked.text = fmtPair(st.linkedSel, st.linkedAll);
            if (UI.embed) UI.embed.text = fmtPair(st.embeddedSel, st.embeddedAll);
            if (UI.broken) UI.broken.text = fmtPair(st.brokenLinkSel, st.brokenLinkAll);

            // Groups
            if (UI.group) UI.group.text = fmtPair(st.groupSel, st.groupAll);
            if (UI.clipGroup) UI.clipGroup.text = fmtPair(st.clipGroupSel, st.clipGroupAll);

            // Paths
            if (UI.pathCount) UI.pathCount.text = fmtPair(st.pathCountSel, st.pathCountAll);
            if (UI.openPath) UI.openPath.text = fmtPair(st.openPathSel, st.openPathAll);
            if (UI.closedPath) UI.closedPath.text = fmtPair(st.closedPathSel, st.closedPathAll);
            if (UI.anchors) UI.anchors.text = fmtPair(st.anchorCountSel, st.anchorCountAll);
            if (UI.handles) UI.handles.text = fmtPair(st.handleCountSel, st.handleCountAll);
            if (UI.compoundPath) UI.compoundPath.text = fmtPair(st.compoundPathSel, st.compoundPathAll);
            if (UI.compoundShape) UI.compoundShape.text = fmtPair(st.compoundShapeSel, st.compoundShapeAll);

            // Guides
            if (UI.rulerGuides) UI.rulerGuides.text = fmtPair(st.rulerGuideSel, st.rulerGuideAll);
            if (UI.artboardGuides) UI.artboardGuides.text = fmtPair(st.artboardGuideSel, st.artboardGuideAll);
            if (UI.otherGuides) UI.otherGuides.text = fmtPair(st.otherGuideSel, st.otherGuideAll);

            try { dlg.layout.layout(true); } catch (e) { }
        }

        var LABEL_WIDTH = 130; /* 全ラベルの幅を統一 / Set uniform label width */
        var LABEL_WIDTH_LEFT = 100; /* 左カラムのラベル幅 / Label width for left column */
        var LABEL_WIDTH_TOP = 90; // オブジェクト・アートボード数専用ラベル幅
        var LABEL_WIDTH_RIGHT = 140; /* 右カラムのラベル幅 / Label width for right column */


        // 2カラムのグループ / Two-column group
        var twoColGroup = dlg.add("group");
        twoColGroup.orientation = "row";
        twoColGroup.alignChildren = ["fill", "top"];

        // 左カラム
        var leftCol = twoColGroup.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        /* その他の詳細を表示するパネル / Panel to display other details */
        var panelOther = leftCol.add("panel", undefined, LABELS.other[lang]);
        panelOther.orientation = "column";
        panelOther.alignChildren = ["fill", "top"];
        panelOther.margins = [15, 20, 0, 10];

        function addOtherRow(label, value) {
            var row = panelOther.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.artboards[lang]) UI.artboards = val;
            if (label === LABELS.objects[lang]) UI.objects = val;
        }

        addOtherRow(LABELS.artboards[lang], String(currentStats.artboardCount));
        addOtherRow(LABELS.objects[lang], currentStats.totalObjSel + " / " + currentStats.totalObjAll);



        /* テキストの詳細を表示するパネル / Panel to display text details */
        var panelText = leftCol.add("panel", undefined, LABELS.texts[lang]);
        panelText.orientation = "column";
        panelText.alignChildren = ["fill", "top"];
        panelText.margins = [15, 20, 0, 10];

        /* 文字・段落の詳細を表示するパネル / Panel to display character and paragraph details */
        var panelCharPara = leftCol.add("panel", undefined, LABELS.charPara[lang]);
        panelCharPara.orientation = "column";
        panelCharPara.alignChildren = ["fill", "top"];
        panelCharPara.margins = [15, 20, 0, 10];


        function addCharParaRow(label, value) {
            var row = panelCharPara.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.chars[lang]) UI.chars = val;
            if (label === LABELS.paras[lang]) UI.paras = val;
            if (label === LABELS.forcedBreaks[lang]) UI.forcedBreaks = val;
        }

        function addTextRow(label, value) {
            var row = panelText.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.texts[lang]) UI.texts = val;
            if (label === LABELS.pointText[lang]) UI.pointText = val;
            if (label === LABELS.areaText[lang]) UI.areaText = val;
            if (label === LABELS.pathText[lang]) UI.pathText = val;
        }

        addCharParaRow(LABELS.chars[lang], currentStats.totalCharSel + " / " + currentStats.totalCharAll);
        addCharParaRow(LABELS.paras[lang], currentStats.paraCountSel + " / " + currentStats.paraCountAll);
        addCharParaRow(LABELS.forcedBreaks[lang], currentStats.forcedBreakSel + " / " + currentStats.forcedBreakAll);
        addTextRow(LABELS.texts[lang], currentStats.textCountSel + " / " + currentStats.textCountAll);
        addTextRow(LABELS.pointText[lang], currentStats.pointTextSel + " / " + currentStats.pointTextAll);
        addTextRow(LABELS.areaText[lang], currentStats.areaTextSel + " / " + currentStats.areaTextAll);
        addTextRow(LABELS.pathText[lang], currentStats.pathTextSel + " / " + currentStats.pathTextAll);

        // 右カラム
        var rightCol = twoColGroup.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];

        /* 透明の詳細を表示するパネル / Panel to display transparency details */
        var panelTransparency = rightCol.add("panel", undefined, LABELS.transparency[lang]);
        panelTransparency.orientation = "column";
        panelTransparency.alignChildren = ["fill", "top"];
        panelTransparency.margins = [15, 20, 0, 10];

        function addTransparencyRow(label, value) {
            var row = panelTransparency.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.opacityLt100[lang]) UI.opacityLt100 = val;
            if (label === LABELS.blendNotNormal[lang]) UI.blendNotNormal = val;
        }

        addTransparencyRow(LABELS.opacityLt100[lang], currentStats.opacityLt100Sel + " / " + currentStats.opacityLt100All);
        addTransparencyRow(LABELS.blendNotNormal[lang], currentStats.blendNotNormalSel + " / " + currentStats.blendNotNormalAll);

        /* 画像の詳細を表示するパネル / Panel to display image details */
        var panelImage = leftCol.add("panel", undefined, LABELS.link[lang]);
        panelImage.orientation = "column";
        panelImage.alignChildren = ["fill", "top"];
        panelImage.margins = [15, 20, 0, 10];

        function addImageRow(label, value) {
            var row = panelImage.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.linked[lang]) UI.linked = val;
            if (label === LABELS.embed[lang]) UI.embed = val;
            if (label === LABELS.broken[lang]) UI.broken = val;
        }

        addImageRow(LABELS.linked[lang], currentStats.linkedSel + " / " + currentStats.linkedAll);
        addImageRow(LABELS.embed[lang], currentStats.embeddedSel + " / " + currentStats.embeddedAll);
        addImageRow(LABELS.broken[lang], currentStats.brokenLinkSel + " / " + currentStats.brokenLinkAll);

        /* グループの詳細を表示するパネル / Panel to display group details */
        var panelGroup = rightCol.add("panel", undefined, LABELS.group[lang]);
        panelGroup.orientation = "column";
        panelGroup.alignChildren = ["fill", "top"];
        panelGroup.margins = [15, 20, 0, 10];

        function addGroupRow(label, value) {
            var row = panelGroup.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.group[lang]) UI.group = val;
            if (label === LABELS.clipGroup[lang]) UI.clipGroup = val;
        }

        addGroupRow(LABELS.group[lang], currentStats.groupSel + " / " + currentStats.groupAll);
        addGroupRow(LABELS.clipGroup[lang], currentStats.clipGroupSel + " / " + currentStats.clipGroupAll);

        /* パスの詳細を表示するパネル / Panel to display path details */
        var panelPath = rightCol.add("panel", undefined, LABELS.path[lang]);
        panelPath.orientation = "column";
        panelPath.alignChildren = ["fill", "top"];
        panelPath.margins = [15, 20, 0, 10];

        /* パス総数表示行を最上部に追加 / Add total path count row at top */
        function addPathRow(label, value) {
            var row = panelPath.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.pathCount[lang]) UI.pathCount = val;
            if (label === LABELS.openPath[lang]) UI.openPath = val;
            if (label === LABELS.closedPath[lang]) UI.closedPath = val;
            if (label === LABELS.anchors[lang]) UI.anchors = val;
            if (label === LABELS.handles[lang]) UI.handles = val;
            if (label === LABELS.compoundPath[lang]) UI.compoundPath = val;
            if (label === LABELS.compoundShape[lang]) UI.compoundShape = val;
        }
        addPathRow(LABELS.pathCount[lang], currentStats.pathCountSel + " / " + currentStats.pathCountAll);
        addPathRow(LABELS.openPath[lang], currentStats.openPathSel + " / " + currentStats.openPathAll);
        addPathRow(LABELS.closedPath[lang], currentStats.closedPathSel + " / " + currentStats.closedPathAll);
        addPathRow(LABELS.anchors[lang], currentStats.anchorCountSel + " / " + currentStats.anchorCountAll);
        addPathRow(LABELS.handles[lang], currentStats.handleCountSel + " / " + currentStats.handleCountAll);
        addPathRow(LABELS.compoundPath[lang], currentStats.compoundPathSel + " / " + currentStats.compoundPathAll);
        addPathRow(LABELS.compoundShape[lang], currentStats.compoundShapeSel + " / " + currentStats.compoundShapeAll);

        /* ガイドの詳細を表示するパネル / Panel to display guide details */
        var panelGuide = rightCol.add("panel", undefined, LABELS.guide[lang]);
        panelGuide.orientation = "column";
        panelGuide.alignChildren = ["fill", "top"];
        panelGuide.margins = [15, 20, 0, 10];

        function addGuideRow(label, value) {
            var row = panelGuide.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
            if (label === LABELS.rulerGuides[lang]) UI.rulerGuides = val;
            if (label === LABELS.artboardGuides[lang]) UI.artboardGuides = val;
            if (label === LABELS.otherGuides[lang]) UI.otherGuides = val;
        }

        addGuideRow(LABELS.rulerGuides[lang], currentStats.rulerGuideSel + " / " + currentStats.rulerGuideAll);
        addGuideRow(LABELS.artboardGuides[lang], currentStats.artboardGuideSel + " / " + currentStats.artboardGuideAll);
        addGuideRow(LABELS.otherGuides[lang], currentStats.otherGuideSel + " / " + currentStats.otherGuideAll);

        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.margins = [10, 10, 10, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        // 左側グループ / Left-side button group
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnExport = btnLeftGroup.add("button", undefined, LABELS.exportPreset[lang], { name: "export" });

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacerL = btnRowGroup.add("statictext", undefined, "");
        spacerL.alignment = ["fill", "fill"];
        spacerL.minimumSize.width = 0;

        // 中央ボタン / Center button
        var btnRefresh = btnRowGroup.add("button", undefined, LABELS.refresh[lang], { name: "refresh" });

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacerR = btnRowGroup.add("statictext", undefined, "");
        spacerR.alignment = ["fill", "fill"];
        spacerR.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnOK = btnRightGroup.add("button", undefined, LABELS.ok[lang], { name: "close" });

        try { dlg.defaultElement = btnRefresh; } catch (e) { }

        btnExport.onClick = function () {
            try {
                __SI_needExport__ = true;
                this.__SI_needExport__ = true;
                __SI_needRefresh__ = true;
                this.__SI_needRefresh__ = true;
            } catch (e) {
                alert((lang === "ja" ? "書き出し中にエラー: " : "Export error: ") + e);
            }
        };
        btnRefresh.onClick = function () {
            try {
                __SI_needRefresh__ = true;
                this.__SI_needRefresh__ = true;
            } catch (e) {
                alert((lang === "ja" ? "更新中にエラー: " : "Refresh error: ") + e);
            }
        };
        btnOK.onClick = function () {
            try {
                if (__SI_idleTask__) {
                    try { __SI_idleTask__.remove(); } catch (eR) { }
                    __SI_idleTask__ = null;
                    this.__SI_idleTask__ = null;
                }
            } catch (e0) { }
            dlg.close();
        };

        // --- Idle handling and export coordination
        function getActiveDocForIdle() {
            // Try to get a usable document handle during idle.
            var n = 0;
            try { n = app.documents.length; } catch (eLen) { n = 0; }
            if (n <= 0) return null;
            try { return app.activeDocument; } catch (e0) { }
            try { return app.documents[0]; } catch (e1) { return null; }
        }

        function getSelectionSafe(doc) {
            var s = [];
            try {
                s = doc.selection;
                if (!s) s = [];
            } catch (e) {
                s = [];
            }
            return s;
        }

        function doExportWithStats(docNow, stats) {
            var docName = "document";
            try { docName = docNow.name.replace(/\.[^\.]+$/, ""); } catch (e) { }
            var today = new Date();
            var yyyy = today.getFullYear();
            var mm = ("0" + (today.getMonth() + 1)).slice(-2);
            var dd = ("0" + today.getDate()).slice(-2);
            var dateStr = yyyy + mm + dd;

            var desktopPath = Folder.desktop + "/count-" + docName + "-" + dateStr + ".txt";
            var file = new File(desktopPath);

            function wPair(labelJa, labelEn, selVal, allVal) {
                if (lang === "ja") file.writeln(labelJa + " " + selVal + " / " + allVal);
                else file.writeln(labelEn + " " + selVal + " / " + allVal);
            }
            function wSingle(labelJa, labelEn, val) {
                if (lang === "ja") file.writeln(labelJa + " " + val);
                else file.writeln(labelEn + " " + val);
            }
            function wSection(titleJa, titleEn) {
                file.writeln("");
                file.writeln(lang === "ja" ? titleJa : titleEn);
            }

            if (!file.open("w")) {
                alert(lang === "ja" ? "ファイルを開けませんでした。" : "Failed to open the file.");
                return;
            }

            // Header
            file.writeln("Selection Inspector Report");
            try { file.writeln("Document: " + docNow.name); } catch (eDoc) { file.writeln("Document: (unknown)"); }
            file.writeln("Date: " + yyyy + "-" + mm + "-" + dd);
            file.writeln("");
            if (lang === "ja") file.writeln("※ 値は『選択 / 全体』の形式です（アートボードのみ全体）");
            else file.writeln("Note: values are formatted as 'Selection / All' (Artboards is All only)");

            // Basics
            wSection("基本", "Basics");
            wSingle(LABELS.artboards.ja.replace(/：$/, ":"), LABELS.artboards.en.replace(/:$/, ":"), stats.artboardCount);
            wPair(LABELS.objects.ja.replace(/：$/, ":"), LABELS.objects.en.replace(/:$/, ":"), stats.totalObjSel, stats.totalObjAll);

            // Text
            wSection("テキスト", "Text Frames");
            wPair(LABELS.texts.ja.replace(/：$/, ":"), LABELS.texts.en.replace(/:$/, ":"), stats.textCountSel, stats.textCountAll);
            wPair(LABELS.pointText.ja.replace(/：$/, ":"), LABELS.pointText.en.replace(/:$/, ":"), stats.pointTextSel, stats.pointTextAll);
            wPair(LABELS.areaText.ja.replace(/：$/, ":"), LABELS.areaText.en.replace(/:$/, ":"), stats.areaTextSel, stats.areaTextAll);
            wPair(LABELS.pathText.ja.replace(/：$/, ":"), LABELS.pathText.en.replace(/:$/, ":"), stats.pathTextSel, stats.pathTextAll);

            // Chars/Paras
            wSection("文字・段落", "Characters & Paragraphs");
            wPair(LABELS.chars.ja.replace(/：$/, ":"), LABELS.chars.en.replace(/:$/, ":"), stats.totalCharSel, stats.totalCharAll);
            wPair(LABELS.paras.ja.replace(/：$/, ":"), LABELS.paras.en.replace(/:$/, ":"), stats.paraCountSel, stats.paraCountAll);
            wPair(LABELS.forcedBreaks.ja.replace(/：$/, ":"), LABELS.forcedBreaks.en.replace(/:$/, ":"), stats.forcedBreakSel, stats.forcedBreakAll);

            // Images
            wSection("配置画像", "Images");
            wPair(LABELS.linked.ja.replace(/：$/, ":"), LABELS.linked.en.replace(/:$/, ":"), stats.linkedSel, stats.linkedAll);
            wPair(LABELS.embed.ja.replace(/：$/, ":"), LABELS.embed.en.replace(/:$/, ":"), stats.embeddedSel, stats.embeddedAll);
            wPair(LABELS.broken.ja.replace(/：$/, ":"), LABELS.broken.en.replace(/:$/, ":"), stats.brokenLinkSel, stats.brokenLinkAll);

            // Transparency
            wSection("透明", "Transparency");
            wPair(LABELS.opacityLt100.ja.replace(/：$/, ":"), LABELS.opacityLt100.en.replace(/:$/, ":"), stats.opacityLt100Sel, stats.opacityLt100All);
            wPair(LABELS.blendNotNormal.ja.replace(/：$/, ":"), LABELS.blendNotNormal.en.replace(/:$/, ":"), stats.blendNotNormalSel, stats.blendNotNormalAll);

            // Groups
            wSection("グループ", "Groups");
            wPair(LABELS.group.ja.replace(/：$/, ":"), LABELS.group.en.replace(/:$/, ":"), stats.groupSel, stats.groupAll);
            wPair(LABELS.clipGroup.ja.replace(/：$/, ":"), LABELS.clipGroup.en.replace(/:$/, ":"), stats.clipGroupSel, stats.clipGroupAll);

            // Paths
            wSection("パス", "Paths");
            wPair(LABELS.pathCount.ja.replace(/：$/, ":"), LABELS.pathCount.en.replace(/:$/, ":"), stats.pathCountSel, stats.pathCountAll);
            wPair(LABELS.openPath.ja.replace(/：$/, ":"), LABELS.openPath.en.replace(/:$/, ":"), stats.openPathSel, stats.openPathAll);
            wPair(LABELS.closedPath.ja.replace(/：$/, ":"), LABELS.closedPath.en.replace(/:$/, ":"), stats.closedPathSel, stats.closedPathAll);
            wPair(LABELS.anchors.ja.replace(/：$/, ":"), LABELS.anchors.en.replace(/:$/, ":"), stats.anchorCountSel, stats.anchorCountAll);
            wPair(LABELS.handles.ja.replace(/：$/, ":"), LABELS.handles.en.replace(/:$/, ":"), stats.handleCountSel, stats.handleCountAll);
            wPair(LABELS.compoundPath.ja.replace(/：$/, ":"), LABELS.compoundPath.en.replace(/:$/, ":"), stats.compoundPathSel, stats.compoundPathAll);
            wPair(LABELS.compoundShape.ja.replace(/：$/, ":"), LABELS.compoundShape.en.replace(/:$/, ":"), stats.compoundShapeSel, stats.compoundShapeAll);

            // Guides
            wSection("ガイド", "Guides");
            wPair(LABELS.rulerGuides.ja.replace(/：$/, ":"), LABELS.rulerGuides.en.replace(/:$/, ":"), stats.rulerGuideSel, stats.rulerGuideAll);
            wPair(LABELS.artboardGuides.ja.replace(/：$/, ":"), LABELS.artboardGuides.en.replace(/:$/, ":"), stats.artboardGuideSel, stats.artboardGuideAll);
            wPair(LABELS.otherGuides.ja.replace(/：$/, ":"), LABELS.otherGuides.en.replace(/:$/, ":"), stats.otherGuideSel, stats.otherGuideAll);

            file.close();
            $.writeln("Exported: " + desktopPath);
        }

        // Setup idle task (runs refresh/export outside palette click context)
        if (app.idleTasks && app.idleTasks.add) {
            // Remove old one if exists
            try {
                if (__SI_idleTask__) {
                    try { __SI_idleTask__.remove(); } catch (eR) { }
                    __SI_idleTask__ = null;
                    this.__SI_idleTask__ = null;
                }
            } catch (e0) { }

            try {
                __SI_idleTask__ = app.idleTasks.add({ name: "SelectionInspectorIdle", sleep: 200 });
                this.__SI_idleTask__ = __SI_idleTask__;
            } catch (e1) {
                __SI_idleTask__ = null;
            }

            if (__SI_idleTask__) {
                __SI_idleTask__.onIdle = function () {
                    try {
                        // Nothing to do
                        if (!__SI_needRefresh__ && !__SI_needExport__) return;

                        var docNow = getActiveDocForIdle();
                        if (!docNow) {
                            __SI_needRefresh__ = false;
                            __SI_needExport__ = false;
                            this.__SI_needRefresh__ = false;
                            this.__SI_needExport__ = false;
                            return;
                        }

                        if (__SI_needRefresh__) {
                            var selNow = getSelectionSafe(docNow);
                            currentStats = computeStats(docNow, selNow);
                            this.__SI_currentStats__ = currentStats;
                            updateUIFromStats(currentStats);
                            __SI_needRefresh__ = false;
                            this.__SI_needRefresh__ = false;
                        }

                        if (__SI_needExport__) {
                            doExportWithStats(docNow, currentStats);
                            __SI_needExport__ = false;
                            this.__SI_needExport__ = false;
                        }
                    } catch (e2) {
                        // Avoid alert storms from idle; log for debugging
                        try { $.writeln("SelectionInspector idle error: " + e2); } catch (e3) { }
                    }
                };
            }
        }

        // Publish context for scheduleTask callbacks
        this.__SI_ctx = {
            computeStats: computeStats,
            updateUIFromStats: updateUIFromStats,
            doExportWithStats: doExportWithStats
        };
        this.__SI_currentStats__ = currentStats;
        this.__SI_lang__ = lang;

        updateUIFromStats(currentStats);

        /* ダイアログ位置と透明度を調整 / Adjust dialog position and opacity */
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);

        dlg.show();
        dlg.active = true;

    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

main();