#target illustrator
#targetengine "PathCleanupToolEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script name:
PathCleanupTool

### 更新日 / Updated:
20260304

### 概要 / Overview:
選択したパス（グループ/複合パス内も含む）を対象に、
- 直線上で冗長なアンカーポイント
- 直線として扱えるベジェ区間上のハンドル
を削除（最適化）します。ロック/非表示（親やレイヤー含む）はスキップします。
許容誤差は「アンカー削除」と「ハンドル削除」で分離しています。

ダイアログ表示時点の選択を固定し、予測（→表示）と実行結果が常に一致するように設計されています。

「プレビュー（別レイヤー）」を有効にすると、実オブジェクトは変更せず、
専用レイヤーに複製を作成して色付け表示し、削除対象のパスを可視化します。
プレビューはOK／キャンセル時に自動的に削除されます。
*/

var SCRIPT_VERSION = "v1.1.7";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "パスの最適化",
        en: "Path Optimization"
    },
    panelProcess: {
        ja: "削除対象",
        en: "Removal Targets"
    },
    panelCurrentInfo: {
        ja: "現在の情報",
        en: "Current Info"
    },
    labelPathCount: {
        ja: "パスの数：",
        en: "Paths: "
    },
    labelAnchorCount: {
        ja: "アンカーポイント数：",
        en: "Anchor points: "
    },
    labelHandleCount: {
        ja: "ハンドル数：",
        en: "Handles: "
    },
    cbRemoveAnchors: {
        ja: "直線上のアンカーポイント",
        en: "Collinear anchor points"
    },
    cbRemoveSameAnchors: {
        ja: "同じ座標のアンカーポイント",
        en: "Duplicate anchor points"
    },
    cbRemoveHandles: {
        ja: "パスと同じ角度のハンドル",
        en: "Handles on straight segments"
    },
    labelTolHandle: {
        ja: "許容誤差",
        en: "Tolerance"
    },
    btnOK: {
        ja: "OK",
        en: "OK"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNeedSelection: {
        ja: "パスを選択してから実行してください。",
        en: "Please select paths before running."
    }
};
    /**
     * 同一座標のアンカーポイント（重複点）を削除（内部：targets 指定）
     * - 連続して同じ座標になっているアンカーのみ対象（離れた位置の同座標は対象外）
     * - オープンパスは端点を削除しない
     * 戻り値：削除したアンカー数
     */
    function removeDuplicateAnchorsOnTargets(targets) {
        if (!targets || !targets.length) return 0;

        var removed = 0;

        for (var s = 0; s < targets.length; s++) {
            var item = targets[s];
            if (!item || isSkippableItem(item)) continue;

            try {
                var pts = item.pathPoints;
                var isClosed = item.closed;
                var n = pts.length;
                if (n < 2) continue;

                // Walk backward so index stays valid after removals
                // Open path: do not remove endpoints => i from n-2 down to 1
                // Closed path: can remove any point, but keep at least 2 points
                var start = isClosed ? (n - 1) : (n - 2);
                var end = isClosed ? 0 : 1;

                for (var i = start; i >= end; i--) {
                    var curLen = pts.length;
                    if (curLen < 2) break;

                    if (i > curLen - 1) i = curLen - 1;
                    if (!isClosed && (i <= 0 || i >= curLen - 1)) continue;

                    var prevIndex = isClosed ? ((i - 1 + curLen) % curLen) : (i - 1);
                    if (prevIndex < 0 || prevIndex > curLen - 1) continue;

                    var pPrev = pts[prevIndex];
                    var pCur = pts[i];

                    if (samePoint(pPrev.anchor, pCur.anchor, TOL_SAMEPOINT)) {
                        pCur.remove();
                        removed++;
                    }
                }
            } catch (_) {
                // ignore
            }
        }

        return removed;
    }

    /**
     * 同一座標のアンカーポイント（重複点）を削除（UI/外部呼び出し用）
     */
    function removeDuplicateAnchors() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeDuplicateAnchorsOnTargets(targets);
    }

function L(key) {
    var v = LABELS[key];
    if (!v) return String(key);
    return v[lang] || v.ja || String(key);
}

(function () {
    // --- shared settings / helpers ---
    /* 共通設定とヘルパー / Shared settings and helpers */
    // Tolerances (keep same default values to avoid behavior change)
    // - Anchor collinearity tolerance: used for redundant anchor removal
    // - Handle collinearity tolerance: used for straight-segment handle normalization
    // - Same-point tolerance: used for handle/anchor coincidence checks and counts
    var TOL_ANCHOR_COLLINEAR = 0.02;
    var TOL_HANDLE_COLLINEAR = 0.01;
    var TOL_SAMEPOINT = 0.02;

    function hasDocument() {
        return app.documents.length > 0;
    }

    function getSelectionOrAlert() {
        if (!hasDocument()) {
            alert(L('alertNoDocument'));
            return null;
        }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!(sel instanceof Array) || sel.length === 0) {
            alert(L('alertNeedSelection'));
            return null;
        }
        return sel;
    }

    // 3点が一直線上にあるか判定（外積を使用）
    function isCollinear(p1, p2, p3, tol) {
        var area = (p2[0] - p1[0]) * (p3[1] - p1[1]) - (p2[1] - p1[1]) * (p3[0] - p1[0]);
        return Math.abs(area) < (tol || TOL_ANCHOR_COLLINEAR);
    }

    function samePoint(a, b, tol) {
        return (Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1])) < (tol || TOL_SAMEPOINT);
    }

    // 点pが線分a-bの直線上にあるか（点から直線までの距離で判定）
    // tol は「距離（pt）」として扱う。線分が極端に短い場合は samePoint にフォールバック。
    function isPointOnLineByDistance(a, b, p, tol) {
        tol = (tol != null) ? tol : TOL_HANDLE_COLLINEAR;

        var abx = b[0] - a[0];
        var aby = b[1] - a[1];
        var len = Math.sqrt(abx * abx + aby * aby);

        if (len < 1e-9) {
            // a と b がほぼ同一点
            return samePoint(a, p, Math.max(TOL_SAMEPOINT, tol));
        }

        // cross product magnitude / |AB| = perpendicular distance
        var apx = p[0] - a[0];
        var apy = p[1] - a[1];
        var area2 = abx * apy - aby * apx; // signed
        var dist = Math.abs(area2) / len;
        return dist <= tol;
    }

    // --- dialog position persistence (session only) ---
    // Store ONLY window location (x,y). Bounds can include size and may restore too small, making UI appear blank.
    // Stored in the target engine's global object; reset when Illustrator quits.
    var DLG_LOC_KEY = "__PathCleanupTool_DialogLocation__";

    function loadDialogLocation() {
        try {
            var p = $.global[DLG_LOC_KEY];
            if (!p || p.length !== 2) return null;
            return [Number(p[0]), Number(p[1])];
        } catch (_) {
            return null;
        }
    }

    function saveDialogLocation(dlg) {
        if (!dlg) return;
        try {
            var p = dlg.location; // [x,y]
            if (!p || p.length !== 2) return;
            $.global[DLG_LOC_KEY] = [Number(p[0]), Number(p[1])];
        } catch (_) {
            // ignore
        }
    }

    function tryRestoreDialogLocation(dlg) {
        var p = loadDialogLocation();
        if (!p) return;
        try {
            dlg.location = p;
        } catch (_) {
            // ignore
        }
    }

    /* ロック・非表示を判定（親・レイヤーも含む） / Check locked/hidden including parents and layers */
    function isSkippableItem(item) {
        var cur = item;
        while (cur) {
            try {
                // PageItem / GroupItem / PathItem etc.
                if (cur.locked === true) return true;
                if (cur.hidden === true) return true;

                // Layer
                if (cur.typename === "Layer") {
                    if (cur.locked === true) return true;
                    if (cur.visible === false) return true;
                }

                // If the item has a layer reference, also respect it
                if (cur.layer) {
                    try {
                        if (cur.layer.locked === true) return true;
                        if (cur.layer.visible === false) return true;
                    } catch (_) { }
                }
            } catch (_) {
                // ignore property access errors
            }

            // Walk up
            try {
                cur = cur.parent;
            } catch (_) {
                break;
            }

            // Stop when reaching document-like root
            if (!cur || cur.typename === "Document") break;
        }
        return false;
    }

    /* 選択からPathItemを再帰的に収集 / Collect PathItems recursively from selection */
    function collectPathItemsFromAny(item, out) {
        if (!item) return;
        if (isSkippableItem(item)) return;

        try {
            // Direct PathItem
            if (item.typename === "PathItem") {
                out.push(item);
                return;
            }

            // CompoundPathItem contains pathItems
            if (item.typename === "CompoundPathItem") {
                for (var i = 0; i < item.pathItems.length; i++) {
                    if (!isSkippableItem(item.pathItems[i])) out.push(item.pathItems[i]);
                }
                return;
            }

            // GroupItem (including clipped groups)
            if (item.typename === "GroupItem") {
                for (var g = 0; g < item.pageItems.length; g++) {
                    collectPathItemsFromAny(item.pageItems[g], out);
                }
                return;
            }

            // Other container types we may want to traverse
            if (item.pageItems && item.pageItems.length) {
                for (var p = 0; p < item.pageItems.length; p++) {
                    collectPathItemsFromAny(item.pageItems[p], out);
                }
            }
        } catch (_) {
            // ignore
        }
    }

    function getTargetPathItemsFromSelection(selection) {
        var out = [];
        for (var i = 0; i < selection.length; i++) {
            collectPathItemsFromAny(selection[i], out);
        }
        return out;
    }

    /* 現在の情報（数）を取得 / Get current info counts */
    function getCurrentInfoCounts() {
        var info = { paths: 0, anchors: 0, handles: 0 };

        if (!hasDocument()) return info;

        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!(sel instanceof Array) || sel.length === 0) return info;

        var targets = getTargetPathItemsFromSelection(sel);

        for (var i = 0; i < targets.length; i++) {
            var item = targets[i];
            if (!item || isSkippableItem(item)) continue;
            info.paths++;

            var pts = item.pathPoints;
            var n = pts.length;
            info.anchors += n;

            for (var k = 0; k < n; k++) {
                var pt = pts[k];
                if (!samePoint(pt.leftDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
                if (!samePoint(pt.rightDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
            }
        }

        return info;
    }

    /* 指定targetsから情報（数）を取得 / Get info counts from given targets */
    function getInfoCountsFromTargets(targets) {
        var info = { paths: 0, anchors: 0, handles: 0 };
        if (!targets || !targets.length) return info;

        for (var i = 0; i < targets.length; i++) {
            var item = targets[i];
            if (!item || isSkippableItem(item)) continue;
            info.paths++;

            var pts = item.pathPoints;
            var n = pts.length;
            info.anchors += n;

            for (var k = 0; k < n; k++) {
                var pt = pts[k];
                if (!samePoint(pt.leftDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
                if (!samePoint(pt.rightDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
            }
        }

        return info;
    }

    /**
     * 不要なアンカーポイント（直線上の冗長点）を削除（内部：targets 指定）
     */
    function removeRedundantAnchorsOnTargets(targets) {
        if (!targets || !targets.length) return 0;

        var removedCount = 0;

        // アンカーポイントからハンドル（方向線）が引き出されていないか判定
        function isStraightPoint(pt) {
            var d1 = Math.abs(pt.anchor[0] - pt.leftDirection[0]) + Math.abs(pt.anchor[1] - pt.leftDirection[1]);
            var d2 = Math.abs(pt.anchor[0] - pt.rightDirection[0]) + Math.abs(pt.anchor[1] - pt.rightDirection[1]);
            return d1 < TOL_SAMEPOINT && d2 < TOL_SAMEPOINT;
        }

        for (var s = 0; s < targets.length; s++) {
            var item = targets[s];
            if (!item || isSkippableItem(item)) continue;

            try {
                var pts = item.pathPoints;
                var isClosed = item.closed;

                // オープンパスの場合は端点を削除しない、クローズドパスの場合は全てループ
                // 削除によってインデックスがずれるのを防ぐため、後ろからループする
                var startIndex = isClosed ? pts.length - 1 : pts.length - 2;
                var endIndex = isClosed ? 0 : 1;

                for (var i = startIndex; i >= endIndex; i--) {
                    var currentLen = pts.length;
                    if (currentLen < 3) break;

                    // 削除でインデックスが範囲外になった場合に備えて補正
                    if (i > currentLen - 1) i = currentLen - 1;
                    if (i < endIndex) break;

                    var prevIndex = (i - 1 + currentLen) % currentLen;
                    var nextIndex = (i + 1) % currentLen;

                    var pA = pts[prevIndex];
                    var pB = pts[i];
                    var pC = pts[nextIndex];

                    if (isStraightPoint(pB) && isCollinear(pA.anchor, pB.anchor, pC.anchor, TOL_ANCHOR_COLLINEAR)) {
                        pB.remove();
                        removedCount++;
                    }
                }
            } catch (_) {
                // ignore error for this PathItem and continue
            }
        }

        return removedCount;
    }

    /**
     * 不要なアンカーポイント（直線上の冗長点）を削除（UI/外部呼び出し用）
     */
    function removeRedundantAnchors() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeRedundantAnchorsOnTargets(targets);
    }

    /**
     * 不要なハンドルを削除（直線になっているベジェ区間のハンドルをアンカーに戻す）（内部：targets 指定）
     * 戻り値：リセットしたハンドル数（左右それぞれ1カウント）
     */
    function removeRedundantHandlesOnTargets(targets) {
        if (!targets || !targets.length) return 0;

        var changed = 0;

        // セグメント(p0 -> p1)が「見た目として直線」か判定
        // 両端アンカーを結ぶ直線に、p0右ハンドル / p1左ハンドルが近ければ直線とみなす
        function isStraightSegment(p0, p1) {
            // 両端アンカーを結ぶ直線に、p0右ハンドル / p1左ハンドルが近ければ直線とみなす
            return isPointOnLineByDistance(p0.anchor, p1.anchor, p0.rightDirection, TOL_HANDLE_COLLINEAR) &&
                isPointOnLineByDistance(p0.anchor, p1.anchor, p1.leftDirection, TOL_HANDLE_COLLINEAR);
        }

        for (var s = 0; s < targets.length; s++) {
            var item = targets[s];
            if (!item || isSkippableItem(item)) continue;

            try {
                var pts = item.pathPoints;
                var len = pts.length;
                if (len < 2) continue;

                var isClosed = item.closed;
                var segCount = isClosed ? len : (len - 1);

                for (var i = 0; i < segCount; i++) {
                    var p0 = pts[i];
                    var p1 = pts[(i + 1) % len];

                    if (!isStraightSegment(p0, p1)) continue;

                    // オープンパス端点のハンドルは触らない / Do not modify endpoint handles for open paths
                    var isOpen = !item.closed;
                    var isFirstSeg = isOpen && (i === 0);
                    var isLastSeg = isOpen && (i === (segCount - 1));

                    // p0.rightDirection: skip if p0 is the first endpoint of an open path
                    if (!isFirstSeg) {
                        if (!samePoint(p0.rightDirection, p0.anchor, TOL_SAMEPOINT)) {
                            p0.rightDirection = p0.anchor;
                            changed++;
                        }
                    }

                    // p1.leftDirection: skip if p1 is the last endpoint of an open path
                    if (!isLastSeg) {
                        if (!samePoint(p1.leftDirection, p1.anchor, TOL_SAMEPOINT)) {
                            p1.leftDirection = p1.anchor;
                            changed++;
                        }
                    }
                }
            } catch (_) {
                // ignore error for this PathItem and continue
            }
        }

        return changed;
    }

    /**
     * 不要なハンドルを削除（直線になっているベジェ区間のハンドルをアンカーに戻す）（UI/外部呼び出し用）
     */
    function removeRedundantHandles() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeRedundantHandlesOnTargets(targets);
    }

    // 予測情報を取得（処理前後のアンカー数とハンドル数）
    // ※実際に削除はせず、指定targetsをスキャンして「実行順（アンカー→ハンドル→アンカー）」をモデル上でシミュレートする
    function getPredictedInfoCountsForTargets(targets, doSameAnchors, doAnchors, doHandles) {
        var infoNow = getInfoCountsFromTargets(targets);

        // --- model helpers ---
        function clonePoint(pt) {
            return {
                a: [pt.anchor[0], pt.anchor[1]],
                l: [pt.leftDirection[0], pt.leftDirection[1]],
                r: [pt.rightDirection[0], pt.rightDirection[1]]
            };
        }

        function clonePath(item) {
            var pts = item.pathPoints;
            var out = [];
            for (var i = 0; i < pts.length; i++) {
                out.push(clonePoint(pts[i]));
            }
            return {
                closed: !!item.closed,
                pts: out
            };
        }

        function isStraightPointM(p) {
            var d1 = Math.abs(p.a[0] - p.l[0]) + Math.abs(p.a[1] - p.l[1]);
            var d2 = Math.abs(p.a[0] - p.r[0]) + Math.abs(p.a[1] - p.r[1]);
            return d1 < TOL_SAMEPOINT && d2 < TOL_SAMEPOINT;
        }

        function samePointM(a, b) {
            return (Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1])) < TOL_SAMEPOINT;
        }

        // アンカー削除（モデル上）
        function removeRedundantAnchorsModel(pathM) {
            var removed = 0;
            var pts = pathM.pts;
            var isClosed = pathM.closed;

            if (pts.length < 3) return 0;

            // 実処理と同じ：オープンは端点を除外、後ろから走査
            var startIndex = isClosed ? (pts.length - 1) : (pts.length - 2);
            var endIndex = isClosed ? 0 : 1;

            for (var i = startIndex; i >= endIndex; i--) {
                var currentLen = pts.length;
                if (currentLen < 3) break;

                // 途中で短くなった場合に備えて補正
                if (i > currentLen - 1) i = currentLen - 1;
                if (i < endIndex) break;

                var prevIndex = (i - 1 + currentLen) % currentLen;
                var nextIndex = (i + 1) % currentLen;

                var pA = pts[prevIndex];
                var pB = pts[i];
                var pC = pts[nextIndex];

                if (isStraightPointM(pB) && isCollinear(pA.a, pB.a, pC.a, TOL_ANCHOR_COLLINEAR)) {
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        // 重複アンカー削除（モデル上）
        function removeDuplicateAnchorsModel(pathM) {
            var removed = 0;
            var pts = pathM.pts;
            var isClosed = pathM.closed;
            var n = pts.length;
            if (n < 2) return 0;

            var start = isClosed ? (n - 1) : (n - 2);
            var end = isClosed ? 0 : 1;

            for (var i = start; i >= end; i--) {
                var curLen = pts.length;
                if (curLen < 2) break;

                if (i > curLen - 1) i = curLen - 1;
                if (!isClosed && (i <= 0 || i >= curLen - 1)) continue;

                var prevIndex = isClosed ? ((i - 1 + curLen) % curLen) : (i - 1);
                if (prevIndex < 0 || prevIndex > curLen - 1) continue;

                var pPrev = pts[prevIndex];
                var pCur = pts[i];

                if (samePointM(pPrev.a, pCur.a)) {
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        // ハンドル削除（モデル上）
        function removeRedundantHandlesModel(pathM) {
            var changed = 0;
            var pts = pathM.pts;
            var len = pts.length;
            if (len < 2) return 0;

            function isStraightSegmentM(p0, p1) {
                return isPointOnLineByDistance(p0.a, p1.a, p0.r, TOL_HANDLE_COLLINEAR) &&
                    isPointOnLineByDistance(p0.a, p1.a, p1.l, TOL_HANDLE_COLLINEAR);
            }

            var isClosed = pathM.closed;
            var segCount = isClosed ? len : (len - 1);

            for (var i = 0; i < segCount; i++) {
                var p0 = pts[i];
                var p1 = pts[(i + 1) % len];

                if (!isStraightSegmentM(p0, p1)) continue;

                var isOpen = !isClosed;
                var isFirstSeg = isOpen && (i === 0);
                var isLastSeg = isOpen && (i === (segCount - 1));

                // p0.rightDirection
                if (!isFirstSeg) {
                    if (!samePointM(p0.r, p0.a)) {
                        p0.r = [p0.a[0], p0.a[1]];
                        changed++;
                    }
                }

                // p1.leftDirection
                if (!isLastSeg) {
                    if (!samePointM(p1.l, p1.a)) {
                        p1.l = [p1.a[0], p1.a[1]];
                        changed++;
                    }
                }
            }

            return changed;
        }

        function countHandlesModel(pathM) {
            var c = 0;
            var pts = pathM.pts;
            for (var i = 0; i < pts.length; i++) {
                var p = pts[i];
                if (!samePointM(p.l, p.a)) c++;
                if (!samePointM(p.r, p.a)) c++;
            }
            return c;
        }

        // --- simulate ---
        var pathsCount = 0;
        var anchorsAfterTotal = 0;
        var handlesAfterTotal = 0;

        for (var t = 0; t < (targets ? targets.length : 0); t++) {
            var item = targets[t];
            if (!item || isSkippableItem(item)) continue;

            pathsCount++;

            var m = clonePath(item);

            // 実行順に合わせる（重複アンカー→冗長アンカー→ハンドル→冗長アンカー→重複アンカー）
            if (doSameAnchors) {
                removeDuplicateAnchorsModel(m);
            }
            if (doAnchors) {
                removeRedundantAnchorsModel(m);
            }
            if (doHandles) {
                removeRedundantHandlesModel(m);
            }
            if (doAnchors) {
                removeRedundantAnchorsModel(m);
            }
            if (doSameAnchors) {
                removeDuplicateAnchorsModel(m);
            }

            anchorsAfterTotal += m.pts.length;
            handlesAfterTotal += countHandlesModel(m);
        }

        return {
            paths: pathsCount,
            anchorsNow: infoNow.anchors,
            anchorsAfter: anchorsAfterTotal,
            handlesNow: infoNow.handles,
            handlesAfter: handlesAfterTotal
        };
    }

    // 互換用：現在の選択を対象に予測 / Backward-compatible wrapper using current selection
    function getPredictedInfoCounts(doSameAnchors, doAnchors, doHandles) {
        if (!hasDocument()) {
            return { paths: 0, anchorsNow: 0, anchorsAfter: 0, handlesNow: 0, handlesAfter: 0 };
        }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!(sel instanceof Array) || sel.length === 0) {
            return { paths: 0, anchorsNow: 0, anchorsAfter: 0, handlesNow: 0, handlesAfter: 0 };
        }
        var targets = getTargetPathItemsFromSelection(sel);
        return getPredictedInfoCountsForTargets(targets, doSameAnchors, doAnchors, doHandles);
    }

    // --- UI ---
    function showDialog(frozenTargets) {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.margins = 18;

        tryRestoreDialogLocation(dlg);
        dlg.onClose = function () {
            saveDialogLocation(dlg);
        };

        var pnlInfo = dlg.add('panel', undefined, L('panelCurrentInfo'));
        pnlInfo.orientation = 'column';
        pnlInfo.alignChildren = ['left', 'top'];
        pnlInfo.margins = [15, 20, 15, 10];

        function addInfoRow(labelKey) {
            var row = pnlInfo.add('group');
            row.orientation = 'row';
            row.alignChildren = ['left', 'center'];

            var stLabel = row.add('statictext', undefined, L(labelKey));
            stLabel.characters = 16;
            stLabel.justify = 'right';

            var stValue = row.add('statictext', undefined, '0');
            stValue.characters = 8;
            return stValue;
        }

        var stPathCount = addInfoRow('labelPathCount');
        var stAnchorCount = addInfoRow('labelAnchorCount');
        var stHandleCount = addInfoRow('labelHandleCount');

        function fmtArrow(a, b) {
            return String(a) + " → " + String(b);
        }

        function refreshInfoPreview(doSameAnchors, doAnchors, doHandles) {
            var p = getPredictedInfoCountsForTargets(frozenTargets, doSameAnchors, doAnchors, doHandles);
            stPathCount.text = String(p.paths);
            stAnchorCount.text = fmtArrow(p.anchorsNow, p.anchorsAfter);
            stHandleCount.text = fmtArrow(p.handlesNow, p.handlesAfter);
        }

        refreshInfoPreview(false, true, true);

        var pnlProcess = dlg.add('panel', undefined, L('panelProcess'));
        pnlProcess.orientation = 'column';
        pnlProcess.alignChildren = ['left', 'top'];
        pnlProcess.margins = [15, 20, 15, 10];

        var cbSameAnchors = pnlProcess.add('checkbox', undefined, L('cbRemoveSameAnchors'));
        cbSameAnchors.value = false;

        var cbAnchors = pnlProcess.add('checkbox', undefined, L('cbRemoveAnchors'));
        cbAnchors.value = true;

        var cbHandle = pnlProcess.add('checkbox', undefined, L('cbRemoveHandles'));
        cbHandle.value = true;

        // Tolerance for straight-segment handle detection (0.01 - 3.00)
        var grpTol = pnlProcess.add('group');
        grpTol.orientation = 'row';
        grpTol.alignChildren = ['left', 'center'];

        var stTol = grpTol.add('statictext', undefined, L('labelTolHandle'));
        stTol.characters = 6;

        var etTol = grpTol.add('edittext', undefined, TOL_HANDLE_COLLINEAR.toFixed(2));
        etTol.characters = 6;

        var slTol = grpTol.add('slider', undefined, Math.round(TOL_HANDLE_COLLINEAR * 100), 1, 300);
        slTol.preferredSize.width = 160;

        function clampTolHandle(v) {
            if (isNaN(v)) return TOL_HANDLE_COLLINEAR;
            if (v < 0.01) v = 0.01;
            if (v > 3) v = 3;
            // keep 2 decimals
            v = Math.round(v * 100) / 100;
            return v;
        }

        function syncTolHandleFromValue(v) {
            v = clampTolHandle(v);
            TOL_HANDLE_COLLINEAR = v;
            etTol.text = v.toFixed(2);
            slTol.value = Math.round(v * 100);
        }

        function parseTolText(s) {
            if (!s) return NaN;
            // accept formats like "[0.01]", "0.01", and full-width brackets
            s = String(s).replace(/\[/g, '').replace(/\]/g, '').replace(/［/g, '').replace(/］/g, '').replace(/\s/g, '');
            return parseFloat(s);
        }

        slTol.onChanging = function () {
            var v = slTol.value / 100;
            syncTolHandleFromValue(v);
            refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        };

        etTol.onChange = function () {
            var v = parseTolText(etTol.text);
            syncTolHandleFromValue(v);
            refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        };

        // init
        syncTolHandleFromValue(TOL_HANDLE_COLLINEAR);

        cbSameAnchors.onClick = function () {
            refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        };
        cbAnchors.onClick = function () {
            refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        };
        cbHandle.onClick = function () {
            grpTol.enabled = cbHandle.value;
            refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        };

        // 初回反映
        refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
        grpTol.enabled = cbHandle.value;

        var btns = dlg.add('group');
        btns.orientation = 'row';
        btns.alignChildren = ['center', 'center'];
        btns.alignment = ['center', 'top'];

        var btnCancel = btns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        var btnOK = btns.add('button', undefined, L('btnOK'), { name: 'ok' });

        btnOK.onClick = function () {
            saveDialogLocation(dlg);
            dlg.close(1);
        };
        btnCancel.onClick = function () {
            saveDialogLocation(dlg);
            dlg.close(0);
        };

        var shown = dlg.show();
        return { ok: (shown === 1), doRemoveSameAnchors: cbSameAnchors.value, doRemoveAnchors: cbAnchors.value, doRemoveHandles: cbHandle.value };
    }

    // ダイアログ表示時点の選択を確定（予測と実行を同一対象に揃える）
    var selectionAtOpen = getSelectionOrAlert();
    if (!selectionAtOpen) return;
    var targetsAtOpen = getTargetPathItemsFromSelection(selectionAtOpen);

    var ui = showDialog(targetsAtOpen);
    if (!ui.ok) return;

    var targetsAtOk = targetsAtOpen;

    var removedAnchors = 0;
    var removedHandles = 0;
    var removedSameAnchors = 0;

    // 実行順：重複アンカー → 冗長アンカー → ハンドル → 冗長アンカー → 重複アンカー
    if (ui.doRemoveSameAnchors) {
        removedSameAnchors += removeDuplicateAnchorsOnTargets(targetsAtOk);
    }
    if (ui.doRemoveAnchors) {
        removedAnchors += removeRedundantAnchorsOnTargets(targetsAtOk);
    }
    if (ui.doRemoveHandles) {
        removedHandles += removeRedundantHandlesOnTargets(targetsAtOk);
    }
    if (ui.doRemoveAnchors) {
        removedAnchors += removeRedundantAnchorsOnTargets(targetsAtOk);
    }
    if (ui.doRemoveSameAnchors) {
        removedSameAnchors += removeDuplicateAnchorsOnTargets(targetsAtOk);
    }

    if (!ui.doRemoveSameAnchors && !ui.doRemoveAnchors && !ui.doRemoveHandles) {
        return;
    }
})();
