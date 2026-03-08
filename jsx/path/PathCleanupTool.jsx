#target illustrator
#targetengine "PathCleanupToolEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script name:
PathCleanupTool

### 更新日 / Updated:
20260309

### 概要 / Overview:
選択したパス（グループ/複合パス内も含む）を対象に、
- 直線上で冗長なアンカーポイント
- 同じ座標のアンカーポイント
- 直線として扱えるベジェ区間上のハンドル
を削除（最適化）します。ロック/非表示（親やレイヤー含む）はスキップします。
さらに「その他」タブでは、スムーズ化／コーナー化／アンカーポイント追加／アンカーポイント分割を実行できます。
許容誤差はハンドル削除用に調整できます。

ダイアログ表示時点の選択を固定し、情報表示と実行対象が一致するようにしています。
その他タブの処理も、OK時に選択固定した同一対象へ確実に適用されるように状態管理を見直しています。
スムーズ化では、前後アンカーが同一点または極端に近い場合でも破綻しにくいようガードを入れています。
また、オープンパスの先頭・末尾は循環参照せず、端点として自然な方向を使うようにしています。
方向は45度刻みに丸めず、前後アンカーから求めた自然な接線方向をそのまま使う「通常のスムーズポイント化」に調整しています。
また、前後セグメントの角度差や長さ差が大きい場合はハンドル長を少し抑えて、鋭角や偏った間隔でも暴れにくいようにしています。
実処理中の例外は、UI系の保存復元とは分けて最小限のログを出すよう整理しています。
*/

var SCRIPT_VERSION = "v1.3";

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
        ja: "情報",
        en: "Info"
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
    },
    tabProcess: {
        ja: "削除対象",
        en: "Removal Targets"
    },
    tabOther: {
        ja: "その他",
        en: "Other"
    },
    rbConvertSmooth: {
        ja: "スムーズポイントに変換",
        en: "Convert to smooth points"
    },
    rbConvertCorner: {
        ja: "コーナーポイントに変換",
        en: "Convert to corner points"
    },
    rbAddAnchors: {
        ja: "アンカーポイントを追加",
        en: "Add anchor points"
    },
    rbSplitAtAnchors: {
        ja: "アンカーポイントで分割",
        en: "Split at anchor points"
    }
};


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

    // 実処理系のみ最小ログ化。UI位置保存や復元系の catch は従来どおり silent のままにする
    function logProcessError(context, e) {
        try {
            $.writeln("[PathCleanupTool] " + context + ": " + e);
        } catch (_) {
            // ignore logging failure
        }
    }

    function hasDocument() {
        return app.documents.length > 0;
    }

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
            } catch (e) {
                logProcessError("removeDuplicateAnchorsOnTargets", e);
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
        } catch (e) {
            logProcessError("collectPathItemsFromAny", e);
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
            } catch (e) {
                logProcessError("removeRedundantAnchorsOnTargets", e);
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
            } catch (e) {
                logProcessError("removeRedundantHandlesOnTargets", e);
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

    // 変換・分割の予測情報を取得
    function getPredictedInfoForConvert(targets, mode) {
        var infoNow = getInfoCountsFromTargets(targets);
        var pathsAfter = infoNow.paths;
        var anchorsAfter = infoNow.anchors;
        var handlesAfter = infoNow.handles;

        if (!targets || !targets.length) {
            return { paths: 0, anchorsNow: 0, anchorsAfter: 0, handlesNow: 0, handlesAfter: 0 };
        }

        if (mode === 'corner') {
            // 全ハンドル削除
            handlesAfter = 0;
        } else if (mode === 'smooth') {
            // 全アンカーに左右ハンドル付与
            handlesAfter = infoNow.anchors * 2;
        } else if (mode === 'add') {
            // 各セグメントに1点追加
            var totalSegs = 0;
            for (var i = 0; i < targets.length; i++) {
                var item = targets[i];
                if (!item || isSkippableItem(item)) continue;
                var n = item.pathPoints.length;
                totalSegs += item.closed ? n : (n - 1);
            }
            anchorsAfter = infoNow.anchors + totalSegs;
            handlesAfter = '-';
        } else if (mode === 'split') {
            // 各セグメントが独立パスに
            var totalSegs2 = 0;
            for (var j = 0; j < targets.length; j++) {
                var item2 = targets[j];
                if (!item2 || isSkippableItem(item2)) continue;
                var n2 = item2.pathPoints.length;
                totalSegs2 += item2.closed ? n2 : (n2 - 1);
            }
            pathsAfter = totalSegs2;
            anchorsAfter = totalSegs2 * 2;
            handlesAfter = '-';
        }

        return {
            paths: infoNow.paths,
            pathsAfter: pathsAfter,
            anchorsNow: infoNow.anchors,
            anchorsAfter: anchorsAfter,
            handlesNow: infoNow.handles,
            handlesAfter: handlesAfter
        };
    }

    // --- UI ---
    function showDialog(frozenTargets) {
        function getActiveMode() {
            return (tpanel.selection === tabOther) ? 'other' : 'process';
        }
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

        function refreshInfoForConvert(mode) {
            var p = getPredictedInfoForConvert(frozenTargets, mode);
            stPathCount.text = fmtArrow(p.paths, p.pathsAfter);
            stAnchorCount.text = fmtArrow(p.anchorsNow, p.anchorsAfter);
            stHandleCount.text = fmtArrow(p.handlesNow, p.handlesAfter);
        }

        function getSelectedConvertMode() {
            if (rbSmooth.value) return 'smooth';
            if (rbCorner.value) return 'corner';
            if (rbAdd.value) return 'add';
            return 'split';
        }

        function refreshByActiveTab() {
            if (tpanel.selection === tabOther) {
                refreshInfoForConvert(getSelectedConvertMode());
            } else {
                refreshInfoPreview(cbSameAnchors.value, cbAnchors.value, cbHandle.value);
            }
        }

        refreshInfoPreview(false, true, true);

        var tpanel = dlg.add('tabbedpanel');
        tpanel.alignChildren = ['fill', 'top'];

        // --- Tab 1: 削除対象 ---
        var tabProcess = tpanel.add('tab', undefined, L('tabProcess'));
        tabProcess.orientation = 'column';
        tabProcess.alignChildren = ['left', 'top'];
        tabProcess.margins = [15, 15, 15, 10];

        var cbSameAnchors = tabProcess.add('checkbox', undefined, L('cbRemoveSameAnchors'));
        cbSameAnchors.value = false;

        var cbAnchors = tabProcess.add('checkbox', undefined, L('cbRemoveAnchors'));
        cbAnchors.value = true;

        var cbHandle = tabProcess.add('checkbox', undefined, L('cbRemoveHandles'));
        cbHandle.value = true;

        // Tolerance for straight-segment handle detection (0.01 - 3.00)
        var grpTol = tabProcess.add('group');
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

        // --- Tab 2: その他 ---
        var tabOther = tpanel.add('tab', undefined, L('tabOther'));
        tabOther.orientation = 'column';
        tabOther.alignChildren = ['left', 'top'];
        tabOther.margins = [15, 15, 15, 10];

        var rbSmooth = tabOther.add('radiobutton', undefined, L('rbConvertSmooth'));
        var rbCorner = tabOther.add('radiobutton', undefined, L('rbConvertCorner'));
        var rbAdd = tabOther.add('radiobutton', undefined, L('rbAddAnchors'));
        var rbSplit = tabOther.add('radiobutton', undefined, L('rbSplitAtAnchors'));
        rbSmooth.value = true;

        rbSmooth.onClick = rbCorner.onClick = rbAdd.onClick = rbSplit.onClick = function () {
            refreshInfoForConvert(getSelectedConvertMode());
        };

        tpanel.onChange = function () {
            refreshByActiveTab();
        };

        tpanel.selection = 0;

        var btns = dlg.add('group');
        btns.orientation = 'row';
        btns.alignChildren = ['center', 'center'];
        btns.alignment = ['center', 'top'];

        var btnCancel = btns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        var btnOK = btns.add('button', undefined, L('btnOK'), { name: 'ok' });

        var result = {
            ok: false,
            activeMode: 'process',
            doRemoveSameAnchors: false,
            doRemoveAnchors: false,
            doRemoveHandles: false,
            convertMode: 'smooth'
        };

        btnOK.onClick = function () {
            result.ok = true;
            result.activeMode = getActiveMode();
            result.doRemoveSameAnchors = cbSameAnchors.value;
            result.doRemoveAnchors = cbAnchors.value;
            result.doRemoveHandles = cbHandle.value;
            result.convertMode = getSelectedConvertMode();
            saveDialogLocation(dlg);
            dlg.close(1);
        };
        btnCancel.onClick = function () {
            saveDialogLocation(dlg);
            dlg.close(0);
        };

        dlg.show();
        return result;
    }

    // --- 変換・分割ロジック ---
    function convertToCorner(pathItem) {
        var points = pathItem.pathPoints;
        for (var k = 0; k < points.length; k++) {
            var pt = points[k];
            pt.pointType = PointType.CORNER;
            pt.leftDirection = pt.anchor;
            pt.rightDirection = pt.anchor;
        }
    }

    function convertToSmooth(pathItem) {
        var points = pathItem.pathPoints;
        for (var k = 0; k < points.length; k++) {
            var pt = points[k];
            pt.pointType = PointType.SMOOTH;

            // クローズドパスは循環参照、オープンパス端点は片側だけを参照する
            var anchor = pt.anchor;
            var isClosed = !!pathItem.closed;
            var prevPt = null;
            var nextPt = null;

            if (isClosed || k > 0) {
                prevPt = points[(k - 1 + points.length) % points.length];
            }
            if (isClosed || k < points.length - 1) {
                nextPt = points[(k + 1) % points.length];
            }

            var dPrev = prevPt ? Math.sqrt(
                Math.pow(prevPt.anchor[0] - anchor[0], 2) +
                Math.pow(prevPt.anchor[1] - anchor[1], 2)
            ) : 0;
            var dNext = nextPt ? Math.sqrt(
                Math.pow(nextPt.anchor[0] - anchor[0], 2) +
                Math.pow(nextPt.anchor[1] - anchor[1], 2)
            ) : 0;

            var lenLeft = dPrev / 3;
            var lenRight = dNext / 3;

            var angleFactor = 1;
            var balanceFactor = 1;

            var prevVecX = prevPt ? (prevPt.anchor[0] - anchor[0]) : 0;
            var prevVecY = prevPt ? (prevPt.anchor[1] - anchor[1]) : 0;
            var nextVecX = nextPt ? (nextPt.anchor[0] - anchor[0]) : 0;
            var nextVecY = nextPt ? (nextPt.anchor[1] - anchor[1]) : 0;

            // 前後アンカーが同一点または極端に近い場合の 0 除算を防ぐ
            var EPS = 1e-6;
            var hasPrev = dPrev > EPS;
            var hasNext = dNext > EPS;

            var dirX = 1;
            var dirY = 0;

            if (hasPrev && hasNext) {
                var tVecX = nextVecX / dNext - prevVecX / dPrev;
                var tVecY = nextVecY / dNext - prevVecY / dPrev;
                var tLen = Math.sqrt(tVecX * tVecX + tVecY * tVecY);

                if (tLen < 0.0001) {
                    dirX = nextVecX / dNext;
                    dirY = nextVecY / dNext;
                } else {
                    dirX = tVecX / tLen;
                    dirY = tVecY / tLen;
                }
            } else if (hasNext) {
                dirX = nextVecX / dNext;
                dirY = nextVecY / dNext;
            } else if (hasPrev) {
                dirX = -prevVecX / dPrev;
                dirY = -prevVecY / dPrev;
            }

            // 鋭角や前後距離の偏りが大きい場合は、ハンドル長を少しだけ抑える
            if (hasPrev && hasNext) {
                var inDirX = -prevVecX / dPrev;
                var inDirY = -prevVecY / dPrev;
                var outDirX = nextVecX / dNext;
                var outDirY = nextVecY / dNext;

                var dot = inDirX * outDirX + inDirY * outDirY;
                if (dot < -1) dot = -1;
                if (dot > 1) dot = 1;

                var turnAngle = Math.acos(dot);
                var angleNorm = turnAngle / Math.PI; // 0..1
                angleFactor = 1 - (0.35 * angleNorm);

                var minLen = Math.min(dPrev, dNext);
                var maxLen = Math.max(dPrev, dNext);
                var imbalance = (maxLen > EPS) ? (1 - (minLen / maxLen)) : 0; // 0..1
                balanceFactor = 1 - (0.25 * imbalance);
            }

            lenLeft *= angleFactor * balanceFactor;
            lenRight *= angleFactor * balanceFactor;

            // 8方向へ丸めず、計算した接線方向をそのまま使う
            pt.leftDirection = [
                anchor[0] - dirX * lenLeft,
                anchor[1] - dirY * lenLeft
            ];
            pt.rightDirection = [
                anchor[0] + dirX * lenRight,
                anchor[1] + dirY * lenRight
            ];
        }
    }

    function convertPathItem(item, mode) {
        var func = (mode === 'smooth') ? convertToSmooth : convertToCorner;
        if (item.typename === 'PathItem') {
            func(item);
        } else if (item.typename === 'CompoundPathItem') {
            for (var i = 0; i < item.pathItems.length; i++) {
                func(item.pathItems[i]);
            }
        } else if (item.typename === 'GroupItem') {
            for (var j = 0; j < item.pageItems.length; j++) {
                convertPathItem(item.pageItems[j], mode);
            }
        }
    }

    function splitAtAnchors(pathItem) {
        var pts = pathItem.pathPoints;
        if (pts.length < 2) return;

        var parent = pathItem.layer;
        var segCount = pathItem.closed ? pts.length : pts.length - 1;

        for (var s = 0; s < segCount; s++) {
            var idx1 = s;
            var idx2 = (s + 1) % pts.length;
            var p1 = pts[idx1];
            var p2 = pts[idx2];

            var newPath = parent.pathItems.add();
            newPath.closed = false;
            newPath.filled = pathItem.filled;
            newPath.fillColor = pathItem.fillColor;
            newPath.stroked = pathItem.stroked;
            if (pathItem.stroked) {
                newPath.strokeColor = pathItem.strokeColor;
                newPath.strokeWidth = pathItem.strokeWidth;
            }

            var np1 = newPath.pathPoints.add();
            np1.anchor = p1.anchor;
            np1.leftDirection = p1.anchor;
            np1.rightDirection = p1.rightDirection;
            np1.pointType = p1.pointType;

            var np2 = newPath.pathPoints.add();
            np2.anchor = p2.anchor;
            np2.leftDirection = p2.leftDirection;
            np2.rightDirection = p2.anchor;
            np2.pointType = p2.pointType;
        }

        pathItem.remove();
    }

    function splitItem(item) {
        if (item.typename === 'PathItem') {
            splitAtAnchors(item);
        } else if (item.typename === 'CompoundPathItem') {
            var paths = [];
            for (var i = 0; i < item.pathItems.length; i++) {
                paths.push(item.pathItems[i]);
            }
            for (var ii = 0; ii < paths.length; ii++) {
                splitAtAnchors(paths[ii]);
            }
        } else if (item.typename === 'GroupItem') {
            var items = [];
            for (var j = 0; j < item.pageItems.length; j++) {
                items.push(item.pageItems[j]);
            }
            for (var jj = 0; jj < items.length; jj++) {
                splitItem(items[jj]);
            }
        }
    }

    // ダイアログ表示時点の選択を確定（予測と実行を同一対象に揃える）
    var selectionAtOpen = getSelectionOrAlert();
    if (!selectionAtOpen) return;
    var targetsAtOpen = getTargetPathItemsFromSelection(selectionAtOpen);

    var ui = showDialog(targetsAtOpen);
    if (!ui.ok) return;

    // タブindex依存だと ScriptUI 環境差で誤判定しうるため、明示状態で分岐する
    if (ui.activeMode === 'other') {
        // --- その他タブ: 変換・分割処理 ---
        var sel = selectionAtOpen.slice(0);
        if (ui.convertMode === 'add') {
            app.executeMenuCommand('Add Anchor Points2');
        } else if (ui.convertMode === 'split') {
            for (var n = 0; n < sel.length; n++) {
                splitItem(sel[n]);
            }
        } else {
            for (var n = 0; n < sel.length; n++) {
                if (!sel[n]) continue;
                convertPathItem(sel[n], ui.convertMode);
            }
        }
    } else {
        // --- 削除対象タブ: クリーンアップ処理 ---
        var targetsAtOk = targetsAtOpen;

        if (!ui.doRemoveSameAnchors && !ui.doRemoveAnchors && !ui.doRemoveHandles) {
            return;
        }

        // 実行順：重複アンカー → 冗長アンカー → ハンドル → 冗長アンカー → 重複アンカー
        if (ui.doRemoveSameAnchors) {
            removeDuplicateAnchorsOnTargets(targetsAtOk);
        }
        if (ui.doRemoveAnchors) {
            removeRedundantAnchorsOnTargets(targetsAtOk);
        }
        if (ui.doRemoveHandles) {
            removeRedundantHandlesOnTargets(targetsAtOk);
        }
        if (ui.doRemoveAnchors) {
            removeRedundantAnchorsOnTargets(targetsAtOk);
        }
        if (ui.doRemoveSameAnchors) {
            removeDuplicateAnchorsOnTargets(targetsAtOk);
        }
    }
})();
