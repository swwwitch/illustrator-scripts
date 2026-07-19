#target illustrator
#targetengine "PathCleanupToolEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要：

- 選択したパス（グループ／複合パスを含む）のパス構造を最適化
- ロック／非表示オブジェクト（親・レイヤー含む）は自動スキップ
- ダイアログ表示時点の選択状態を固定し、情報表示と実行対象を一致

### 主な機能：

- 直線上の冗長なアンカーポイントの削除
- 同一座標のアンカーポイントの削除
- 直線として扱えるベジェ区間のハンドルの整理
- 「その他」タブを3グループ（変換：スムーズ／コーナー、追加：中間／極点、その他：分割／マド埋め）に整理して実行可能
- マド埋めは選択にマド（複合パス）が無いとき自動的に無効化
- アンカー削除用／ハンドル削除用の許容誤差を個別調整可能
- スムーズ化は前後アンカーの距離・角度に応じて安定するよう補正
- オープンパス端点は安全に処理（循環参照なし）
- 実行中の例外は最小限ログ出力

### 紹介記事（note）

https://note.com/dtp_tranist/n/nd82f59bf63a8

### 更新履歴：

- v1.5.2 (2026-07-14) : 「極点を追加」モードを追加（現行版）
- v1.5.1 (2026-03-20)
- v1.0 (2026-03-01) : 初版

*/

/*

### Script Name:

PathCleanupTool.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathCleanupTool.md

### Description:

- Optimizes the structure of selected paths (including inside groups / compound paths)
- Locked / hidden objects (including parents and layers) are skipped automatically
- Selection is locked at dialog open time so the display and execution targets always match

### Main Features:

- Removes redundant anchors on straight segments
- Removes duplicate-coordinate anchors
- Cleans up handles on bezier segments that can be treated as straight
- "Other" tab is organized into 3 groups (Convert: smooth / corner, Add: midpoints / extrema, Other: split / fill holes)
- Fill holes is disabled automatically when the selection has no compound path
- Independent tolerance controls for anchor removal and handle removal
- Smoothing is corrected for stability based on adjacent anchor distance / angle
- Open-path endpoints are processed safely (no circular references)
- Minimal logging for runtime exceptions

### Changelog:

- v1.5.2 (2026-07-14) : Added "Add Extreme Points" mode (current version)
- v1.5.1 (2026-03-20)
- v1.0 (2026-03-01) : Initial release

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathCleanupTool";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-14";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* 許容誤差の初期値（ダイアログで調整可能） / Default tolerances (adjustable in dialog) */
var TOL_ANCHOR_COLLINEAR = 0.02; /* 直線上アンカー削除の許容誤差 / Redundant-anchor removal */
var TOL_HANDLE_COLLINEAR = 0.01; /* 直線区間ハンドル整理の許容誤差 / Straight-segment handle normalization */
var TOL_SAMEPOINT = 0.02;        /* 同一点判定の許容誤差 / Coincident-point check */

// =========================================
// ローカライズ / Localization
// =========================================

/** 現在の UI 言語（"ja" / "en"）を判定して返します。 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/**
 * 日英ラベル定義（カテゴリ別） / Japanese-English label definitions (by category)
 * L("dialog.title") のようにドット区切りで参照する。短い文言は1行で記述。
 * @type {Object}
 */
var LABELS = {
    dialog: {
        title: { ja: "パスの最適化", en: "Path Optimization" }
    },
    panel: {
        info:          { ja: "情報", en: "Info" },
        convertPoints: { ja: "アンカーポイントを変換", en: "Convert anchor points" },
        addPoints:     { ja: "アンカーポイントを追加", en: "Add anchor points" },
        pathOps:       { ja: "その他", en: "Other" }
    },
    tab: {
        process: { ja: "削除対象", en: "Removal Targets" },
        other:   { ja: "変換", en: "Transform" }
    },
    checkbox: {
        removeSameAnchors: { ja: "同じ座標のアンカーポイント", en: "Duplicate anchor points" },
        removeAnchors:     { ja: "直線上のアンカーポイント", en: "Collinear anchor points" },
        removeHandles:     { ja: "パスと同じ角度のハンドル", en: "Handles on straight segments" }
    },
    radio: {
        convertSmooth:  { ja: "スムーズポイントに", en: "To smooth points" },
        convertCorner:  { ja: "コーナーポイントに", en: "To corner points" },
        addAnchors:     { ja: "中間に追加", en: "At midpoints" },
        addExtremePoints: { ja: "極点を追加", en: "Add Extreme Points" },
        splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at anchor points" },
        fillHoles:      { ja: "マド埋め", en: "Fill holes" }
    },
    label: {
        pathCount:   { ja: "パスの数", en: "Paths" },
        anchorCount: { ja: "アンカーポイント数", en: "Anchor points" },
        handleCount: { ja: "ハンドル数", en: "Handles" },
        tolAnchor:   { ja: "許容誤差", en: "Tolerance" },
        tolHandle:   { ja: "許容誤差", en: "Tolerance" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        needSelection: { ja: "パスを選択してから実行してください。", en: "Please select paths before running." }
    },
    tooltip: {
        removeSameAnchors: { ja: "連続して同じ座標にあるアンカーポイントを1つに統合します（離れた位置の同座標は対象外）。", en: "Merges consecutive anchors that share the same coordinates (non-adjacent duplicates are ignored)." },
        removeAnchors:     { ja: "前後のアンカーと一直線上にある冗長なアンカーポイントを削除します。", en: "Removes redundant anchors that lie on a straight line between their neighbors." },
        removeHandles:     { ja: "直線とみなせる区間のハンドルをアンカーに戻します（見た目を変えずにハンドルを整理）。", en: "Resets handles on segments that are effectively straight (tidies handles without changing appearance)." },
        tolAnchor:         { ja: "値が大きいほど、より緩く「直線上」と判定して多くのアンカーを削除します（0.01〜3.00）。", en: "Higher values treat more anchors as collinear and remove more of them (0.01–3.00)." },
        tolHandle:         { ja: "値が大きいほど、より緩く「直線」と判定して多くのハンドルを戻します（0.01〜3.00）。", en: "Higher values treat more segments as straight and reset more handles (0.01–3.00)." },
        addAnchors:        { ja: "各セグメントの中間に1点ずつアンカーポイントを追加します。", en: "Adds one anchor at the midpoint of each segment." },
        addExtremePoints:  { ja: "曲線の上下左右の端（水平・垂直の接線位置＝極点）にアンカーポイントを追加します。", en: "Adds anchors at the curve's extrema (points of horizontal/vertical tangency)." },
        splitAtAnchors:    { ja: "各セグメントを独立したオープンパスに分割します（元のパスは削除）。", en: "Splits each segment into a separate open path (the original path is removed)." },
        fillHoles:         { ja: "複合パスを解除して合体し、穴（マド）を埋めます。選択に複合パスが無い場合は使用できません。", en: "Releases the compound path and unites it to fill holes. Unavailable when the selection has no compound path." }
    }
};

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;               /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウに共通レイアウトを適用します。
 * @param {Window} win - 対象のダイアログウィンドウ。
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）。
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネル（タブを含む）に共通レイアウトを適用します。
 * @param {Panel} panel - 対象のパネル。
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）。
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループ（ボタン列など）に共通レイアウトを適用します。
 * ボタン類をパネル幅いっぱいに広げないため alignChildren は設定しない。
 * @param {Group} group - 対象の行グループ。
 * @param {string} [alignment] - 水平方向の揃え（省略時は "left"）。
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）。
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/** ボタン button の高さを px 詰めます（レイアウト確定後に呼ぶ）。 */
function trimButtonHeight(button, px) {
    try {
        button.size = [button.size.width, button.size.height - px];
    } catch (e) {}
}


(function () {
    // --- shared settings / helpers ---
    /* 共通設定とヘルパー / Shared settings and helpers */
    /* 許容誤差は冒頭「ユーザー設定 / User Settings」で定義（ダイアログで上書き）
       Tolerances are defined in the "User Settings" section at the top (overridden by the dialog). */

    /**
     * パス情報の集計値。
     * @typedef {Object} InfoCounts
     * @property {number} paths - パス数。
     * @property {number} anchors - アンカーポイント数。
     * @property {number} handles - ハンドル数。
     */

    /**
     * シミュレーション用のパスモデル（DOMを書き換えず増減を予測する軽量表現）。
     * @typedef {Object} PathModel
     * @property {boolean} closed - クローズドパスなら true。
     * @property {Array<PathModelPoint>} pts - アンカー点の配列。
     */

    /**
     * パスモデルの1点。各座標は [x, y] の数値配列。
     * @typedef {Object} PathModelPoint
     * @property {Array<number>} a - アンカー座標。
     * @property {Array<number>} l - 左方向線（leftDirection）座標。
     * @property {Array<number>} r - 右方向線（rightDirection）座標。
     */

    /**
     * 実処理中の例外を最小限ログ出力します（UI位置の保存・復元系は従来どおり silent）。
     * @param {string} context - エラー発生箇所を示すラベル。
     * @param {Error} e - 捕捉した例外。
     * @returns {void}
     */
    function logProcessError(context, e) {
        try {
            $.writeln("[PathCleanupTool] " + context + ": " + e);
        } catch (_) {
            // ignore logging failure
        }
    }

    /** ドキュメントが1つ以上開かれていれば true。 */
    function hasDocument() {
        return app.documents.length > 0;
    }

    /**
     * 同一座標のアンカーポイント（重複点）を削除します（対象 targets を明示指定）。
     * 連続して同座標のアンカーのみ対象（離れた位置の同座標は対象外）。オープンパスの端点は削除しない。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @returns {number} 削除したアンカー数。
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
                    if (curLen <= 2) break;

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
     * 互換用：現在の選択を対象に重複アンカーを削除します（後方互換ラッパー）。
     * 現行フローはダイアログ表示時点の targets を固定して removeDuplicateAnchorsOnTargets() を直接呼ぶ。
     * @returns {number} 削除したアンカー数。
     */
    function removeDuplicateAnchors() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeDuplicateAnchorsOnTargets(targets);
    }

    /**
     * ネストした LABELS 定義から現在の言語のラベルを取得します。
     * @param {string} keyPath - "dialog.title" のようなドット区切りのキー。
     * @returns {string} 現在の言語のラベル文字列。該当なしの場合はキー文字列。
     */
    function L(keyPath) {
        var parts = String(keyPath).split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!node) break;
            node = node[parts[i]];
        }
        if (!node) return String(keyPath);
        return node[lang] || node.ja || String(keyPath);
    }

    /** ラベル keyPath 末尾にコロンを付けて返します（日本語は全角、英語は半角）。 */
    function labelText(keyPath) {
        return L(keyPath) + (lang === 'ja' ? '：' : ': ');
    }

    /**
     * ドキュメント・選択の有無を検証し、選択配列を返します。無ければ警告を表示。
     * @returns {?Array<PageItem>} 選択オブジェクト配列。ドキュメント無し／未選択なら null。
     */
    function getSelectionOrAlert() {
        if (!hasDocument()) {
            alert(L('alert.noDocument'));
            return null;
        }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!(sel instanceof Array) || sel.length === 0) {
            alert(L('alert.needSelection'));
            return null;
        }
        return sel;
    }

    /**
     * 3点が一直線上にあるかを外積で判定します。
     * @param {Array<number>} pointA - 点A [x, y]。
     * @param {Array<number>} pointB - 点B [x, y]。
     * @param {Array<number>} pointC - 点C [x, y]。
     * @param {number} [tolerance] - 許容誤差（省略時は TOL_ANCHOR_COLLINEAR）。
     * @returns {boolean} 一直線上とみなせれば true。
     */
    function isCollinear(pointA, pointB, pointC, tolerance) {
        var area = (pointB[0] - pointA[0]) * (pointC[1] - pointA[1]) - (pointB[1] - pointA[1]) * (pointC[0] - pointA[0]);
        return Math.abs(area) < (tolerance || TOL_ANCHOR_COLLINEAR);
    }

    /**
     * 2点がほぼ同一座標かを判定します（マンハッタン距離で比較）。
     * @param {Array<number>} pointA - 点A [x, y]。
     * @param {Array<number>} pointB - 点B [x, y]。
     * @param {number} [tolerance] - 許容誤差（省略時は TOL_SAMEPOINT）。
     * @returns {boolean} 同一とみなせれば true。
     */
    function samePoint(pointA, pointB, tolerance) {
        return (Math.abs(pointA[0] - pointB[0]) + Math.abs(pointA[1] - pointB[1])) < (tolerance || TOL_SAMEPOINT);
    }

    /**
     * 点が線分の延長を含む直線上にあるかを、点から直線までの垂直距離で判定します。
     * 線分が極端に短い場合は samePoint にフォールバックします。
     * @param {Array<number>} lineStart - 直線の始点 [x, y]。
     * @param {Array<number>} lineEnd - 直線の終点 [x, y]。
     * @param {Array<number>} testPoint - 判定する点 [x, y]。
     * @param {number} [tolerance] - 距離の許容誤差（pt、省略時は TOL_HANDLE_COLLINEAR）。
     * @returns {boolean} 直線上とみなせれば true。
     */
    function isPointOnLineByDistance(lineStart, lineEnd, testPoint, tolerance) {
        tolerance = (tolerance != null) ? tolerance : TOL_HANDLE_COLLINEAR;

        var abx = lineEnd[0] - lineStart[0];
        var aby = lineEnd[1] - lineStart[1];
        var len = Math.sqrt(abx * abx + aby * aby);

        if (len < 1e-9) {
            // a と b がほぼ同一点
            return samePoint(lineStart, testPoint, Math.max(TOL_SAMEPOINT, tolerance));
        }

        // cross product magnitude / |AB| = perpendicular distance
        var apx = testPoint[0] - lineStart[0];
        var apy = testPoint[1] - lineStart[1];
        var area2 = abx * apy - aby * apx; // signed
        var dist = Math.abs(area2) / len;
        return dist <= tolerance;
    }

    // --- dialog position persistence (session only) ---
    // Store ONLY window location (x,y). Bounds can include size and may restore too small, making UI appear blank.
    // Stored in the target engine's global object; reset when Illustrator quits.
    var DLG_LOC_KEY = "__PathCleanupTool_DialogLocation__";

    /**
     * 保存済みのダイアログ表示位置を読み込みます（セッション内のみ）。
     * @returns {?Array<number>} [x, y] 位置。未保存・不正なら null。
     */
    function loadDialogLocation() {
        try {
            var savedLocation = $.global[DLG_LOC_KEY];
            if (!savedLocation || savedLocation.length !== 2) return null;
            return [Number(savedLocation[0]), Number(savedLocation[1])];
        } catch (_) {
            return null;
        }
    }

    /**
     * ダイアログの表示位置を保存します（セッション内のみ）。
     * @param {Window} dlg - 対象のダイアログウィンドウ。
     * @returns {void}
     */
    function saveDialogLocation(dlg) {
        if (!dlg) return;
        try {
            var dialogLocation = dlg.location; // [x,y]
            if (!dialogLocation || dialogLocation.length !== 2) return;
            $.global[DLG_LOC_KEY] = [Number(dialogLocation[0]), Number(dialogLocation[1])];
        } catch (_) {
            // ignore
        }
    }

    /**
     * 保存済み位置があればダイアログに復元します。
     * @param {Window} dlg - 対象のダイアログウィンドウ。
     * @returns {void}
     */
    function tryRestoreDialogLocation(dlg) {
        var savedLocation = loadDialogLocation();
        if (!savedLocation) return;
        try {
            dlg.location = savedLocation;
        } catch (_) {
            // ignore
        }
    }

    /**
     * オブジェクトがロック／非表示か（親・レイヤーを遡って）を判定します。
     * @param {PageItem} item - 判定対象のオブジェクト。
     * @returns {boolean} ロックまたは非表示なら true。
     */
    function isSkippableItem(item) {
        var currentItem = item;
        while (currentItem) {
            try {
                // PageItem / GroupItem / PathItem etc.
                if (currentItem.locked === true) return true;
                if (currentItem.hidden === true) return true;

                // Layer
                if (currentItem.typename === "Layer") {
                    if (currentItem.locked === true) return true;
                    if (currentItem.visible === false) return true;
                }

                // If the item has a layer reference, also respect it
                if (currentItem.layer) {
                    try {
                        if (currentItem.layer.locked === true) return true;
                        if (currentItem.layer.visible === false) return true;
                    } catch (_) { }
                }
            } catch (_) {
                // ignore property access errors
            }

            // Walk up
            try {
                currentItem = currentItem.parent;
            } catch (_) {
                break;
            }

            // Stop when reaching document-like root
            if (!currentItem || currentItem.typename === "Document") break;
        }
        return false;
    }

    /**
     * 選択オブジェクトから PathItem を再帰的に収集します（グループ／複合パスを展開）。
     * @param {PageItem} item - 走査対象のオブジェクト。
     * @param {Array<PathItem>} pathItems - 収集先の配列（破壊的に追加）。
     * @returns {void}
     */
    function collectPathItemsFromAny(item, pathItems) {
        if (!item) return;
        if (isSkippableItem(item)) return;

        try {
            // Direct PathItem
            if (item.typename === "PathItem") {
                pathItems.push(item);
                return;
            }

            // CompoundPathItem contains pathItems
            if (item.typename === "CompoundPathItem") {
                for (var i = 0; i < item.pathItems.length; i++) {
                    if (!isSkippableItem(item.pathItems[i])) pathItems.push(item.pathItems[i]);
                }
                return;
            }

            // GroupItem (including clipped groups)
            if (item.typename === "GroupItem") {
                for (var g = 0; g < item.pageItems.length; g++) {
                    collectPathItemsFromAny(item.pageItems[g], pathItems);
                }
                return;
            }

            // Other container types we may want to traverse
            if (item.pageItems && item.pageItems.length) {
                for (var p = 0; p < item.pageItems.length; p++) {
                    collectPathItemsFromAny(item.pageItems[p], pathItems);
                }
            }
        } catch (e) {
            logProcessError("collectPathItemsFromAny", e);
        }
    }

    /**
     * 選択配列から処理対象の PathItem 一覧を収集して返します。
     * @param {Array<PageItem>} selection - 選択オブジェクト配列。
     * @returns {Array<PathItem>} 対象 PathItem 配列。
     */
    function getTargetPathItemsFromSelection(selection) {
        var pathItems = [];
        for (var i = 0; i < selection.length; i++) {
            collectPathItemsFromAny(selection[i], pathItems);
        }
        return pathItems;
    }

    /**
     * オブジェクト（およびグループ内）に、マド（穴）を持つ複合パスが含まれるかを再帰判定します。
     * サブパスが2つ以上の CompoundPathItem を「マドあり」とみなします。ロック／非表示は対象外。
     * @param {PageItem} item - 判定対象のオブジェクト。
     * @returns {boolean} マドを持つ複合パスがあれば true。
     */
    function itemHasHoles(item) {
        if (!item || isSkippableItem(item)) return false;
        try {
            if (item.typename === 'CompoundPathItem') {
                return item.pathItems.length >= 2;
            }
            if (item.typename === 'GroupItem') {
                for (var i = 0; i < item.pageItems.length; i++) {
                    if (itemHasHoles(item.pageItems[i])) return true;
                }
            }
        } catch (e) {
            logProcessError('itemHasHoles', e);
        }
        return false;
    }

    /**
     * 選択内にマド（穴）を持つ複合パスが1つでもあるかを判定します（マド埋めの有効判定に使用）。
     * @param {Array<PageItem>} selection - 選択オブジェクト配列。
     * @returns {boolean} マドを持つ複合パスがあれば true。
     */
    function selectionHasHoles(selection) {
        if (!selection || !selection.length) return false;
        for (var i = 0; i < selection.length; i++) {
            if (itemHasHoles(selection[i])) return true;
        }
        return false;
    }

    /**
     * 互換用：現在の選択から情報（数）を取得します（後方互換ラッパー）。
     * 現行の情報表示はダイアログ表示時点の targets を固定して getInfoCountsFromTargets() を使う。
     * @returns {InfoCounts} パス数・アンカー数・ハンドル数。
     */
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

    /**
     * 指定 targets からパス数・アンカー数・ハンドル数を集計します。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @returns {InfoCounts} 集計結果。
     */
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
     * 直線上の冗長なアンカーポイントを削除します（対象 targets を明示指定）。
     * ハンドルが引き出されておらず前後アンカーと一直線上にある点のみ削除。オープンパスの端点は削除しない。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @returns {number} 削除したアンカー数。
     */
    function removeRedundantAnchorsOnTargets(targets) {
        if (!targets || !targets.length) return 0;

        var removedCount = 0;

        /**
         * アンカーからハンドル（方向線）が引き出されていない（直線的な点）かを判定します。
         * @param {PathPoint} pt - 判定対象のアンカーポイント。
         * @returns {boolean} 左右ハンドルともアンカーと一致すれば true。
         */
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
     * 互換用：現在の選択を対象に直線上の冗長アンカーを削除します（後方互換ラッパー）。
     * @returns {number} 削除したアンカー数。
     */
    function removeRedundantAnchors() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeRedundantAnchorsOnTargets(targets);
    }

    /**
     * 直線になっているベジェ区間のハンドルをアンカーに戻します（対象 targets を明示指定）。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @returns {number} リセットしたハンドル数（左右それぞれ1カウント）。
     */
    function removeRedundantHandlesOnTargets(targets) {
        if (!targets || !targets.length) return 0;

        var changed = 0;

        /**
         * セグメント (p0 → p1) が見た目として直線かを判定します。
         * 両端アンカーを結ぶ直線に p0 右ハンドル・p1 左ハンドルが近ければ直線とみなす。
         * @param {PathPoint} p0 - 始点アンカー。
         * @param {PathPoint} p1 - 終点アンカー。
         * @returns {boolean} 直線とみなせれば true。
         */
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
     * 互換用：現在の選択を対象に直線区間上の冗長ハンドルを削除します（後方互換ラッパー）。
     * @returns {number} リセットしたハンドル数。
     */
    function removeRedundantHandles() {
        var selection = getSelectionOrAlert();
        if (!selection) return 0;
        var targets = getTargetPathItemsFromSelection(selection);
        return removeRedundantHandlesOnTargets(targets);
    }

    /**
     * クリーンアップ実行後のアンカー数・ハンドル数を予測します。
     * DOM は書き換えず、指定 targets を複製したモデル上で実行順（重複→冗長→ハンドル→冗長→重複）をシミュレートする。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @param {boolean} doSameAnchors - 重複アンカー削除を含める場合は true。
     * @param {boolean} doAnchors - 直線上の冗長アンカー削除を含める場合は true。
     * @param {boolean} doHandles - 直線区間のハンドル削除を含める場合は true。
     * @returns {{paths: number, anchorsNow: number, anchorsAfter: number, handlesNow: number, handlesAfter: number}} 予測結果。
     */
    function getPredictedInfoCountsForTargets(targets, doSameAnchors, doAnchors, doHandles) {
        var infoNow = getInfoCountsFromTargets(targets);

        // --- model helpers ---
        /** PathPoint pt をモデル点 PathModelPoint（{a, l, r}）に複製します。 */
        function clonePoint(pt) {
            return {
                a: [pt.anchor[0], pt.anchor[1]],
                l: [pt.leftDirection[0], pt.leftDirection[1]],
                r: [pt.rightDirection[0], pt.rightDirection[1]]
            };
        }

        /**
         * PathItem を編集可能なパスモデルに複製します。
         * @param {PathItem} item - 複製元のパス。
         * @returns {PathModel} 複製したパスモデル。
         */
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

        /** モデル点 p にハンドルが無い（直線的：左右ともアンカーと一致）なら true。 */
        function isStraightPointModel(p) {
            var d1 = Math.abs(p.a[0] - p.l[0]) + Math.abs(p.a[1] - p.l[1]);
            var d2 = Math.abs(p.a[0] - p.r[0]) + Math.abs(p.a[1] - p.r[1]);
            return d1 < TOL_SAMEPOINT && d2 < TOL_SAMEPOINT;
        }

        /** 2つのモデル座標 a, b がほぼ同一（TOL_SAMEPOINT 基準）なら true。 */
        function samePointModel(a, b) {
            return (Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1])) < TOL_SAMEPOINT;
        }

        /**
         * モデル上で直線上の冗長アンカーを削除します。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} 削除したアンカー数。
         */
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

                if (isStraightPointModel(pB) && isCollinear(pA.a, pB.a, pC.a, TOL_ANCHOR_COLLINEAR)) {
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        /**
         * モデル上で重複（同一座標）アンカーを削除します。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} 削除したアンカー数。
         */
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
                if (curLen <= 2) break;

                if (i > curLen - 1) i = curLen - 1;
                if (!isClosed && (i <= 0 || i >= curLen - 1)) continue;

                var prevIndex = isClosed ? ((i - 1 + curLen) % curLen) : (i - 1);
                if (prevIndex < 0 || prevIndex > curLen - 1) continue;

                var pPrev = pts[prevIndex];
                var pCur = pts[i];

                if (samePointModel(pPrev.a, pCur.a)) {
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        /**
         * モデル上で直線区間のハンドルをアンカーに戻します。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} リセットしたハンドル数。
         */
        function removeRedundantHandlesModel(pathM) {
            var changed = 0;
            var pts = pathM.pts;
            var len = pts.length;
            if (len < 2) return 0;

            /**
             * モデル上でセグメント (p0 → p1) が直線かを判定します。
             * @param {PathModelPoint} p0 - 始点モデル点。
             * @param {PathModelPoint} p1 - 終点モデル点。
             * @returns {boolean} 直線とみなせれば true。
             */
            function isStraightSegmentModel(p0, p1) {
                return isPointOnLineByDistance(p0.a, p1.a, p0.r, TOL_HANDLE_COLLINEAR) &&
                    isPointOnLineByDistance(p0.a, p1.a, p1.l, TOL_HANDLE_COLLINEAR);
            }

            var isClosed = pathM.closed;
            var segCount = isClosed ? len : (len - 1);

            for (var i = 0; i < segCount; i++) {
                var p0 = pts[i];
                var p1 = pts[(i + 1) % len];

                if (!isStraightSegmentModel(p0, p1)) continue;

                var isOpen = !isClosed;
                var isFirstSeg = isOpen && (i === 0);
                var isLastSeg = isOpen && (i === (segCount - 1));

                // p0.rightDirection
                if (!isFirstSeg) {
                    if (!samePointModel(p0.r, p0.a)) {
                        p0.r = [p0.a[0], p0.a[1]];
                        changed++;
                    }
                }

                // p1.leftDirection
                if (!isLastSeg) {
                    if (!samePointModel(p1.l, p1.a)) {
                        p1.l = [p1.a[0], p1.a[1]];
                        changed++;
                    }
                }
            }

            return changed;
        }

        /**
         * モデル上のハンドル数を数えます（左右それぞれ1カウント）。
         * @param {PathModel} pathM - 対象のパスモデル。
         * @returns {number} ハンドル数。
         */
        function countHandlesModel(pathM) {
            var c = 0;
            var pts = pathM.pts;
            for (var i = 0; i < pts.length; i++) {
                var p = pts[i];
                if (!samePointModel(p.l, p.a)) c++;
                if (!samePointModel(p.r, p.a)) c++;
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

            var pathModel = clonePath(item);

            // 実行順に合わせる（重複アンカー→冗長アンカー→ハンドル→冗長アンカー→重複アンカー）
            if (doSameAnchors) {
                removeDuplicateAnchorsModel(pathModel);
            }
            if (doAnchors) {
                removeRedundantAnchorsModel(pathModel);
            }
            if (doHandles) {
                removeRedundantHandlesModel(pathModel);
            }
            if (doAnchors) {
                removeRedundantAnchorsModel(pathModel);
            }
            if (doSameAnchors) {
                removeDuplicateAnchorsModel(pathModel);
            }

            anchorsAfterTotal += pathModel.pts.length;
            handlesAfterTotal += countHandlesModel(pathModel);
        }

        return {
            paths: pathsCount,
            anchorsNow: infoNow.anchors,
            anchorsAfter: anchorsAfterTotal,
            handlesNow: infoNow.handles,
            handlesAfter: handlesAfterTotal
        };
    }

    /**
     * 互換用：現在の選択を対象に予測情報を取得します（後方互換ラッパー）。
     * @param {boolean} doSameAnchors - 重複アンカー削除を含める場合は true。
     * @param {boolean} doAnchors - 直線上の冗長アンカー削除を含める場合は true。
     * @param {boolean} doHandles - 直線区間のハンドル削除を含める場合は true。
     * @returns {{paths: number, anchorsNow: number, anchorsAfter: number, handlesNow: number, handlesAfter: number}} 予測結果。
     */
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

    /**
     * 「その他」タブ（変換・分割）の予測情報を取得します。
     * corner=全ハンドル削除／smooth=全アンカーに左右ハンドル付与／add=各セグメントに1点追加／
     * split=各セグメントを独立パス化／fillHoles=結果を事前算出できないため未確定（"-"）。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @param {string} mode - 変換モード（'smooth' | 'corner' | 'add' | 'split' | 'fillHoles'）。
     * @returns {{paths: number, pathsAfter: (number|string), anchorsNow: number, anchorsAfter: (number|string), handlesNow: number, handlesAfter: (number|string)}} 予測結果。
     */
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
        } else if (mode === 'extreme') {
            // 各曲線セグメントの極点数を実測して加算
            anchorsAfter = infoNow.anchors + countExtremaForTargets(targets);
            handlesAfter = '-';
        } else if (mode === 'fillHoles') {
            // パスファインダー結果は事前に正確に出せないため未確定
            pathsAfter = '-';
            anchorsAfter = '-';
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
    /**
     * showDialog() の戻り値。
     * @typedef {Object} DialogResult
     * @property {boolean} ok - OK で閉じたら true。
     * @property {string} activeMode - 実行タブ（'process' | 'other'）。
     * @property {boolean} doRemoveSameAnchors - 重複アンカー削除を行うか。
     * @property {boolean} doRemoveAnchors - 直線上の冗長アンカー削除を行うか。
     * @property {boolean} doRemoveHandles - 直線区間のハンドル削除を行うか。
     * @property {string} convertMode - 変換モード（'smooth' | 'corner' | 'add' | 'split' | 'fillHoles'）。
     */

    /**
     * メインダイアログを表示し、ユーザーの選択結果を返します。
     * @param {Array<PathItem>} frozenTargets - ダイアログ表示時点で固定した対象パス（予測表示に使用）。
     * @param {boolean} hasHoles - 選択内にマド（穴）を持つ複合パスがあるか（false のとき「マド埋め」を無効化）。
     * @returns {DialogResult} 実行内容を表す結果オブジェクト。
     */
    function showDialog(frozenTargets, hasHoles) {
        /**
         * 現在アクティブなタブに対応する実行モードを返します。
         * @returns {string} 'other'（その他タブ）または 'process'（削除対象タブ）。
         */
        function getActiveMode() {
            return (tabbedPanel.selection === tabOther) ? 'other' : 'process';
        }
        var dlg = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dlg);

        tryRestoreDialogLocation(dlg);
        dlg.onClose = function () {
            saveDialogLocation(dlg);
        };

        var infoPanel = dlg.add('panel', undefined, L('panel.info'));
        setupPanel(infoPanel);

        /**
         * 情報パネルに「ラベル：値」の行を追加します。
         * @param {string} labelKey - ラベルの LABELS ドット区切りキー。
         * @returns {StaticText} 値表示用の StaticText（後で text を更新する）。
         */
        function addInfoRow(labelKey) {
            var infoRow = infoPanel.add('group');
            infoRow.orientation = 'row';
            infoRow.alignChildren = ['left', 'center'];

            var rowLabel = infoRow.add('statictext', undefined, labelText(labelKey));
            rowLabel.characters = 13;
            rowLabel.justify = 'right';

            var rowValue = infoRow.add('statictext', undefined, '0');
            rowValue.characters = 4;
            return rowValue;
        }

        var pathCountValue = addInfoRow('label.pathCount');
        var anchorCountValue = addInfoRow('label.anchorCount');
        var handleCountValue = addInfoRow('label.handleCount');

        /** 「beforeValue → afterValue」の矢印連結文字列を生成します。 */
        function formatArrow(beforeValue, afterValue) {
            return String(beforeValue) + " → " + String(afterValue);
        }

        /**
         * 「削除対象」タブの予測値で情報パネルを更新します。
         * @param {boolean} doSameAnchors - 重複アンカー削除を含めるか。
         * @param {boolean} doAnchors - 直線上の冗長アンカー削除を含めるか。
         * @param {boolean} doHandles - 直線区間のハンドル削除を含めるか。
         * @returns {void}
         */
        function refreshInfoPreview(doSameAnchors, doAnchors, doHandles) {
            var predictedInfo = getPredictedInfoCountsForTargets(frozenTargets, doSameAnchors, doAnchors, doHandles);
            pathCountValue.text = String(predictedInfo.paths);
            anchorCountValue.text = formatArrow(predictedInfo.anchorsNow, predictedInfo.anchorsAfter);
            handleCountValue.text = formatArrow(predictedInfo.handlesNow, predictedInfo.handlesAfter);
        }

        /**
         * 「その他」タブの予測値で情報パネルを更新します。
         * @param {string} mode - 変換モード（'smooth' | 'corner' | 'add' | 'split' | 'fillHoles'）。
         * @returns {void}
         */
        function refreshInfoForConvert(mode) {
            var predictedInfo = getPredictedInfoForConvert(frozenTargets, mode);
            pathCountValue.text = formatArrow(predictedInfo.paths, predictedInfo.pathsAfter);
            anchorCountValue.text = formatArrow(predictedInfo.anchorsNow, predictedInfo.anchorsAfter);
            handleCountValue.text = formatArrow(predictedInfo.handlesNow, predictedInfo.handlesAfter);
        }

        /**
         * 「その他」タブで選択中のラジオボタンから変換モードを返します。
         * @returns {string} 'smooth' | 'corner' | 'add' | 'split' | 'fillHoles'。
         */
        function getSelectedConvertMode() {
            if (smoothRadio.value) return 'smooth';
            if (cornerRadio.value) return 'corner';
            if (addAnchorsRadio.value) return 'add';
            if (extremePointsRadio.value) return 'extreme';
            if (splitRadio.value) return 'split';
            return 'fillHoles';
        }

        /**
         * アクティブなタブに応じて情報パネルの予測表示を切り替えます。
         * @returns {void}
         */
        function refreshByActiveTab() {
            if (tabbedPanel.selection === tabOther) {
                refreshInfoForConvert(getSelectedConvertMode());
            } else {
                refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
            }
        }

        refreshInfoPreview(false, true, true);

        var tabbedPanel = dlg.add('tabbedpanel');
        tabbedPanel.alignChildren = ['fill', 'top'];

        // --- Tab 1: 削除対象 ---
        var tabProcess = tabbedPanel.add('tab', undefined, L('tab.process'));
        setupPanel(tabProcess);
        /* タブ内の右余白を詰める（PANEL_MARGINS の右16→6）。サブプロパティ代入は反映されないため配列で上書き */
        tabProcess.margins = [16, 20, 6, 12];

        var removeSameAnchorsCheckbox = tabProcess.add('checkbox', undefined, L('checkbox.removeSameAnchors'));
        removeSameAnchorsCheckbox.helpTip = L('tooltip.removeSameAnchors');
        removeSameAnchorsCheckbox.value = true;

        // Tolerance for collinear anchor detection (0.01 - 3.00)

        var anchorOptionsGroup = tabProcess.add('group');
        anchorOptionsGroup.orientation = 'column';
        anchorOptionsGroup.alignChildren = ['left', 'top'];
        anchorOptionsGroup.margins = [0, 15, 0, 15];

        var removeAnchorsCheckbox = anchorOptionsGroup.add('checkbox', undefined, L('checkbox.removeAnchors'));
        removeAnchorsCheckbox.helpTip = L('tooltip.removeAnchors');
        removeAnchorsCheckbox.value = true;

        var anchorToleranceRow = anchorOptionsGroup.add('group');
        anchorToleranceRow.orientation = 'row';
        anchorToleranceRow.alignChildren = ['left', 'center'];
        anchorToleranceRow.margins = [20, 0, 0, 0];

        var anchorToleranceLabel = anchorToleranceRow.add('statictext', undefined, L('label.tolAnchor'));
        anchorToleranceLabel.helpTip = L('tooltip.tolAnchor');
        anchorToleranceLabel.characters = 6;

        var anchorToleranceInput = anchorToleranceRow.add('edittext', undefined, TOL_ANCHOR_COLLINEAR.toFixed(2));
        anchorToleranceInput.helpTip = L('tooltip.tolAnchor');
        anchorToleranceInput.characters = 6;

        var anchorToleranceSlider = anchorOptionsGroup.add('slider', undefined, Math.round(TOL_ANCHOR_COLLINEAR * 100), 1, 300);
        anchorToleranceSlider.helpTip = L('tooltip.tolAnchor');
        anchorToleranceSlider.preferredSize.width = 160;
        anchorToleranceSlider.indent = 20;

        /**
         * アンカー許容誤差を有効範囲（0.01〜3.00、小数2桁）に丸めます。
         * @param {number} toleranceValue - 入力値。
         * @returns {number} 丸めた許容誤差。NaN の場合は現在値。
         */
        function clampAnchorTolerance(toleranceValue) {
            if (isNaN(toleranceValue)) return TOL_ANCHOR_COLLINEAR;
            if (toleranceValue < 0.01) toleranceValue = 0.01;
            if (toleranceValue > 3) toleranceValue = 3;
            toleranceValue = Math.round(toleranceValue * 100) / 100;
            return toleranceValue;
        }

        /**
         * アンカー許容誤差を確定し、入力欄・スライダー・内部値を同期します。
         * @param {number} toleranceValue - 設定する許容誤差。
         * @returns {void}
         */
        function syncAnchorToleranceFromValue(toleranceValue) {
            toleranceValue = clampAnchorTolerance(toleranceValue);
            TOL_ANCHOR_COLLINEAR = toleranceValue;
            anchorToleranceInput.text = toleranceValue.toFixed(2);
            anchorToleranceSlider.value = Math.round(toleranceValue * 100);
        }

        anchorToleranceSlider.onChanging = function () {
            var toleranceValue = anchorToleranceSlider.value / 100;
            syncAnchorToleranceFromValue(toleranceValue);
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };

        anchorToleranceInput.onChange = function () {
            var toleranceValue = parseToleranceText(anchorToleranceInput.text);
            syncAnchorToleranceFromValue(toleranceValue);
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };

        syncAnchorToleranceFromValue(TOL_ANCHOR_COLLINEAR);

        var handleOptionsGroup = tabProcess.add('group');
        handleOptionsGroup.orientation = 'column';
        handleOptionsGroup.alignChildren = ['left', 'top'];
        // handleOptionsGroup.margins = [0, 0, 0, 8];

        var removeHandlesCheckbox = handleOptionsGroup.add('checkbox', undefined, L('checkbox.removeHandles'));
        removeHandlesCheckbox.helpTip = L('tooltip.removeHandles');
        removeHandlesCheckbox.value = true;

        // Tolerance for straight-segment handle detection (0.01 - 3.00)
        var handleToleranceRow = handleOptionsGroup.add('group');
        handleToleranceRow.orientation = 'row';
        handleToleranceRow.alignChildren = ['left', 'center'];
        handleToleranceRow.margins = [20, 0, 0, 0];

        var handleToleranceLabel = handleToleranceRow.add('statictext', undefined, L('label.tolHandle'));
        handleToleranceLabel.helpTip = L('tooltip.tolHandle');
        handleToleranceLabel.characters = 6;

        var handleToleranceInput = handleToleranceRow.add('edittext', undefined, TOL_HANDLE_COLLINEAR.toFixed(2));
        handleToleranceInput.helpTip = L('tooltip.tolHandle');
        handleToleranceInput.characters = 6;

        var handleToleranceSlider = handleOptionsGroup.add('slider', undefined, Math.round(TOL_HANDLE_COLLINEAR * 100), 1, 300);
        handleToleranceSlider.helpTip = L('tooltip.tolHandle');
        handleToleranceSlider.preferredSize.width = 160;
        handleToleranceSlider.indent = 20;

        /**
         * ハンドル許容誤差を有効範囲（0.01〜3.00、小数2桁）に丸めます。
         * @param {number} toleranceValue - 入力値。
         * @returns {number} 丸めた許容誤差。NaN の場合は現在値。
         */
        function clampHandleTolerance(toleranceValue) {
            if (isNaN(toleranceValue)) return TOL_HANDLE_COLLINEAR;
            if (toleranceValue < 0.01) toleranceValue = 0.01;
            if (toleranceValue > 3) toleranceValue = 3;
            // keep 2 decimals
            toleranceValue = Math.round(toleranceValue * 100) / 100;
            return toleranceValue;
        }

        /**
         * ハンドル許容誤差を確定し、入力欄・スライダー・内部値を同期します。
         * @param {number} toleranceValue - 設定する許容誤差。
         * @returns {void}
         */
        function syncHandleToleranceFromValue(toleranceValue) {
            toleranceValue = clampHandleTolerance(toleranceValue);
            TOL_HANDLE_COLLINEAR = toleranceValue;
            handleToleranceInput.text = toleranceValue.toFixed(2);
            handleToleranceSlider.value = Math.round(toleranceValue * 100);
        }

        /**
         * 許容誤差の入力文字列を数値に変換します（角括弧・全角括弧・空白を除去）。
         * @param {string} text - 入力文字列。
         * @returns {number} 変換した数値。空文字なら NaN。
         */
        function parseToleranceText(text) {
            if (!text) return NaN;
            // accept formats like "[0.01]", "0.01", and full-width brackets
            text = String(text).replace(/\[/g, '').replace(/\]/g, '').replace(/［/g, '').replace(/］/g, '').replace(/\s/g, '');
            return parseFloat(text);
        }

        handleToleranceSlider.onChanging = function () {
            var toleranceValue = handleToleranceSlider.value / 100;
            syncHandleToleranceFromValue(toleranceValue);
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };

        handleToleranceInput.onChange = function () {
            var toleranceValue = parseToleranceText(handleToleranceInput.text);
            syncHandleToleranceFromValue(toleranceValue);
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };

        // init
        syncHandleToleranceFromValue(TOL_HANDLE_COLLINEAR);

        removeSameAnchorsCheckbox.onClick = function () {
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };
        removeAnchorsCheckbox.onClick = function () {
            anchorToleranceRow.enabled = removeAnchorsCheckbox.value;
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };
        removeHandlesCheckbox.onClick = function () {
            handleToleranceRow.enabled = removeHandlesCheckbox.value;
            refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        };

        // 初回反映
        refreshInfoPreview(removeSameAnchorsCheckbox.value, removeAnchorsCheckbox.value, removeHandlesCheckbox.value);
        anchorToleranceRow.enabled = removeAnchorsCheckbox.value;
        handleToleranceRow.enabled = removeHandlesCheckbox.value;

        // --- Tab 2: その他 ---
        var tabOther = tabbedPanel.add('tab', undefined, L('tab.other'));
        setupPanel(tabOther);
        /* タブ内の右余白を詰める（PANEL_MARGINS の右16→6）。サブプロパティ代入は反映されないため配列で上書き */
        tabOther.margins = [16, 20, 6, 12];

        // パネル1：アンカーポイントを変換（スムーズ／コーナー）
        var convertPointsPanel = tabOther.add('panel', undefined, L('panel.convertPoints'));
        setupPanel(convertPointsPanel);
        var smoothRadio = convertPointsPanel.add('radiobutton', undefined, L('radio.convertSmooth'));
        var cornerRadio = convertPointsPanel.add('radiobutton', undefined, L('radio.convertCorner'));

        // パネル2：アンカーポイントを追加（アンカー追加／極点追加）
        var addPointsPanel = tabOther.add('panel', undefined, L('panel.addPoints'));
        setupPanel(addPointsPanel);
        var addAnchorsRadio = addPointsPanel.add('radiobutton', undefined, L('radio.addAnchors'));
        addAnchorsRadio.helpTip = L('tooltip.addAnchors');
        var extremePointsRadio = addPointsPanel.add('radiobutton', undefined, L('radio.addExtremePoints'));
        extremePointsRadio.helpTip = L('tooltip.addExtremePoints');

        // パネル3：その他（分割／マド埋め）
        var pathOpsPanel = tabOther.add('panel', undefined, L('panel.pathOps'));
        setupPanel(pathOpsPanel);
        var splitRadio = pathOpsPanel.add('radiobutton', undefined, L('radio.splitAtAnchors'));
        splitRadio.helpTip = L('tooltip.splitAtAnchors');
        var fillHolesRadio = pathOpsPanel.add('radiobutton', undefined, L('radio.fillHoles'));
        fillHolesRadio.helpTip = L('tooltip.fillHoles');

        // 選択内にマド（複合パス）が無ければ「マド埋め」は無効化（ディム表示）
        if (!hasHoles) {
            fillHolesRadio.enabled = false;
        }

        smoothRadio.value = true;

        // パネルをまたぐ radiobutton は ScriptUI の自動排他が効かないため、
        // 全ラジオを1グループとして手動で単一選択を維持する
        var convertRadios = [smoothRadio, cornerRadio, addAnchorsRadio, extremePointsRadio, splitRadio, fillHolesRadio];

        /**
         * 指定ラジオだけを選択状態にし、他をすべて解除して予測表示を更新します。
         * @param {RadioButton} selectedRadio - 選択状態にするラジオボタン。
         * @returns {void}
         */
        function selectConvertRadio(selectedRadio) {
            for (var r = 0; r < convertRadios.length; r++) {
                convertRadios[r].value = (convertRadios[r] === selectedRadio);
            }
            refreshInfoForConvert(getSelectedConvertMode());
        }

        for (var radioIndex = 0; radioIndex < convertRadios.length; radioIndex++) {
            (function (radio) {
                radio.onClick = function () {
                    selectConvertRadio(radio);
                };
            })(convertRadios[radioIndex]);
        }

        tabbedPanel.onChange = function () {
            refreshByActiveTab();
        };

        tabbedPanel.selection = 0;

        var buttonRow = dlg.add('group');
        buttonRow.orientation = 'row';
        buttonRow.alignChildren = ['center', 'center'];
        buttonRow.alignment = ['center', 'top'];
        buttonRow.margins = [0, 10, 0, 0];

        var btnCancel = buttonRow.add('button', undefined, L('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.add('button', undefined, L('button.ok'), { name: 'ok' });

        var dialogResult = {
            ok: false,
            activeMode: 'process',
            doRemoveSameAnchors: false,
            doRemoveAnchors: false,
            doRemoveHandles: false,
            convertMode: 'smooth'
        };

        btnOK.onClick = function () {
            dialogResult.ok = true;
            dialogResult.activeMode = getActiveMode();
            dialogResult.doRemoveSameAnchors = removeSameAnchorsCheckbox.value;
            dialogResult.doRemoveAnchors = removeAnchorsCheckbox.value;
            dialogResult.doRemoveHandles = removeHandlesCheckbox.value;
            dialogResult.convertMode = getSelectedConvertMode();
            saveDialogLocation(dlg);
            dlg.close(1);
        };
        btnCancel.onClick = function () {
            saveDialogLocation(dlg);
            dlg.close(0);
        };

        dlg.show();
        return dialogResult;
    }

    // --- 変換・分割ロジック ---
    /**
     * パスの全アンカーをコーナーポイント化し、ハンドルを除去します。
     * @param {PathItem} pathItem - 対象のパス。
     * @returns {void}
     */
    function convertToCorner(pathItem) {
        var points = pathItem.pathPoints;
        for (var k = 0; k < points.length; k++) {
            var pt = points[k];
            pt.pointType = PointType.CORNER;
            pt.leftDirection = pt.anchor;
            pt.rightDirection = pt.anchor;
        }
    }

    /**
     * パスの全アンカーをスムーズポイント化し、前後アンカーの距離・角度に応じてハンドルを補正します。
     * @param {PathItem} pathItem - 対象のパス。
     * @returns {void}
     */
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

    /**
     * オブジェクトを再帰的にたどり、含まれるパスへスムーズ／コーナー変換を適用します。
     * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
     * @param {string} mode - 'smooth' または 'corner'。
     * @returns {void}
     */
    function convertPathItem(item, mode) {
        if (!item || isSkippableItem(item)) return;
        var convertFn = (mode === 'smooth') ? convertToSmooth : convertToCorner;
        if (item.typename === 'PathItem') {
            convertFn(item);
        } else if (item.typename === 'CompoundPathItem') {
            for (var i = 0; i < item.pathItems.length; i++) {
                if (!isSkippableItem(item.pathItems[i])) {
                    convertFn(item.pathItems[i]);
                }
            }
        } else if (item.typename === 'GroupItem') {
            for (var j = 0; j < item.pageItems.length; j++) {
                if (!isSkippableItem(item.pageItems[j])) {
                    convertPathItem(item.pageItems[j], mode);
                }
            }
        }
    }

    /**
     * パスを各セグメントごとの独立したオープンパスに分割します（元のパスは削除）。
     * @param {PathItem} pathItem - 分割対象のパス。
     * @returns {void}
     */
    function splitAtAnchors(pathItem) {
        var pts = pathItem.pathPoints;
        if (pts.length < 2) return;

        var parentLayer = pathItem.layer;
        var segCount = pathItem.closed ? pts.length : pts.length - 1;

        for (var s = 0; s < segCount; s++) {
            var startIndex = s;
            var endIndex = (s + 1) % pts.length;
            var startPoint = pts[startIndex];
            var endPoint = pts[endIndex];

            var newPath = parentLayer.pathItems.add();
            newPath.closed = false;
            newPath.filled = pathItem.filled;
            newPath.fillColor = pathItem.fillColor;
            newPath.stroked = pathItem.stroked;
            if (pathItem.stroked) {
                newPath.strokeColor = pathItem.strokeColor;
                newPath.strokeWidth = pathItem.strokeWidth;
            }

            var newStartPoint = newPath.pathPoints.add();
            newStartPoint.anchor = startPoint.anchor;
            newStartPoint.leftDirection = startPoint.anchor;
            newStartPoint.rightDirection = startPoint.rightDirection;
            newStartPoint.pointType = startPoint.pointType;

            var newEndPoint = newPath.pathPoints.add();
            newEndPoint.anchor = endPoint.anchor;
            newEndPoint.leftDirection = endPoint.leftDirection;
            newEndPoint.rightDirection = endPoint.anchor;
            newEndPoint.pointType = endPoint.pointType;
        }

        pathItem.remove();
    }

    /**
     * マド埋め：複合パス解除 → ライブパスファインダー（合体）→ アピアランス展開 → グループ解除を実行します。
     * @returns {void}
     */
    function fillHolesOnSelection() {
        app.executeMenuCommand('group');
        app.executeMenuCommand('noCompoundPath');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
        try {
            app.executeMenuCommand('ungroup');
        } catch (e) {
            // 単一オブジェクトでグループ化されていない場合があるため握りつぶす
        }
    }

    /**
     * オブジェクトを再帰的にたどり、含まれるパスをアンカーで分割します。
     * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
     * @returns {void}
     */
    function splitItem(item) {
        if (!item || isSkippableItem(item)) return;
        if (item.typename === 'PathItem') {
            splitAtAnchors(item);
        } else if (item.typename === 'CompoundPathItem') {
            var childPaths = [];
            for (var i = 0; i < item.pathItems.length; i++) {
                if (!isSkippableItem(item.pathItems[i])) {
                    childPaths.push(item.pathItems[i]);
                }
            }
            for (var ii = 0; ii < childPaths.length; ii++) {
                splitAtAnchors(childPaths[ii]);
            }
        } else if (item.typename === 'GroupItem') {
            var childItems = [];
            for (var j = 0; j < item.pageItems.length; j++) {
                if (!isSkippableItem(item.pageItems[j])) {
                    childItems.push(item.pageItems[j]);
                }
            }
            for (var jj = 0; jj < childItems.length; jj++) {
                splitItem(childItems[jj]);
            }
        }
    }

    // --- 極点（Extreme Points）追加ロジック ---
    /** 2点 a, b を t（0〜1）で線形補間した座標 [x, y] を返します。 */
    function lerpPoint(a, b, t) {
        return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
    }

    /**
     * 3次ベジェを媒介変数 t で2本に分割します（de Casteljau）。
     * @param {Array<Array<number>>} controlPoints - 制御点4つ [P0, P1, P2, P3]。
     * @param {number} t - 分割位置（0〜1）。
     * @returns {{left: Array<Array<number>>, right: Array<Array<number>>}} 分割後の左右の制御点列。
     */
    function splitCubic(controlPoints, t) {
        var p01 = lerpPoint(controlPoints[0], controlPoints[1], t);
        var p12 = lerpPoint(controlPoints[1], controlPoints[2], t);
        var p23 = lerpPoint(controlPoints[2], controlPoints[3], t);
        var p012 = lerpPoint(p01, p12, t);
        var p123 = lerpPoint(p12, p23, t);
        var p0123 = lerpPoint(p012, p123, t);
        return { left: [controlPoints[0], p01, p012, p0123], right: [p0123, p123, p23, controlPoints[3]] };
    }

    /**
     * 1次元の3次ベジェ f(t) について f'(t)=0 となる t（極値）を求めます。
     * @param {number} v0 - P0 成分。
     * @param {number} v1 - P1 成分。
     * @param {number} v2 - P2 成分。
     * @param {number} v3 - P3 成分。
     * @returns {Array<number>} (0,1) の範囲にある t の配列。
     */
    function extremaParamsOfComponent(v0, v1, v2, v3) {
        var diff0 = v1 - v0;
        var diff1 = v2 - v1;
        var diff2 = v3 - v2;
        var quadA = diff0 - 2 * diff1 + diff2;
        var quadB = 2 * (diff1 - diff0);
        var quadC = diff0;
        var roots = [];
        var EPS = 1e-6;
        if (Math.abs(quadA) < EPS) {
            if (Math.abs(quadB) > EPS) roots.push(-quadC / quadB);
        } else {
            var discriminant = quadB * quadB - 4 * quadA * quadC;
            if (discriminant >= 0) {
                var sqrtDiscriminant = Math.sqrt(discriminant);
                roots.push((-quadB + sqrtDiscriminant) / (2 * quadA));
                roots.push((-quadB - sqrtDiscriminant) / (2 * quadA));
            }
        }
        var paramsInRange = [];
        for (var i = 0; i < roots.length; i++) {
            if (roots[i] > EPS && roots[i] < 1 - EPS) paramsInRange.push(roots[i]);
        }
        return paramsInRange;
    }

    /**
     * セグメント（p0→p1）でハンドルが水平/垂直になる媒介変数 t を求めます。
     * 直線セグメント（両端ハンドルがアンカーと一致）は対象外で空配列を返す。
     * @param {PathPoint} p0 - 始点アンカー。
     * @param {PathPoint} p1 - 終点アンカー。
     * @returns {Array<number>} 昇順・重複除去済みの t 配列。
     */
    function bezierExtremaParams(p0, p1) {
        var isStraight = samePoint(p0.rightDirection, p0.anchor, TOL_SAMEPOINT) &&
            samePoint(p1.leftDirection, p1.anchor, TOL_SAMEPOINT);
        if (isStraight) return [];

        var P0 = p0.anchor, P1 = p0.rightDirection, P2 = p1.leftDirection, P3 = p1.anchor;
        var paramsX = extremaParamsOfComponent(P0[0], P1[0], P2[0], P3[0]); /* 垂直接線 dx/dt=0 */
        var paramsY = extremaParamsOfComponent(P0[1], P1[1], P2[1], P3[1]); /* 水平接線 dy/dt=0 */
        var extremaParams = paramsX.concat(paramsY);

        extremaParams.sort(function (x, y) { return x - y; });
        var uniqueParams = [];
        for (var i = 0; i < extremaParams.length; i++) {
            if (uniqueParams.length === 0 || Math.abs(extremaParams[i] - uniqueParams[uniqueParams.length - 1]) > 1e-4) uniqueParams.push(extremaParams[i]);
        }
        return uniqueParams;
    }

    /**
     * セグメントを複数の t（元セグメントの媒介変数）で分割し、
     * 端点ハンドルの更新値と挿入する内部アンカーを返します。
     * @param {Array<number>} P0 - 始点アンカー座標。
     * @param {Array<number>} P1 - 始点右ハンドル座標。
     * @param {Array<number>} P2 - 終点左ハンドル座標。
     * @param {Array<number>} P3 - 終点アンカー座標。
     * @param {Array<number>} extremaParams - 昇順の分割 t 配列。
     * @returns {{startRight: Array<number>, endLeft: Array<number>, interior: Array<Object>}} 分割結果。
     */
    function splitSegmentAtParams(P0, P1, P2, P3, extremaParams) {
        var currentCurve = [P0, P1, P2, P3];
        var prevParam = 0;
        var interior = [];
        var startRight = P1;
        for (var k = 0; k < extremaParams.length; k++) {
            var localT = (extremaParams[k] - prevParam) / (1 - prevParam);
            var pieces = splitCubic(currentCurve, localT);
            if (k === 0) {
                startRight = pieces.left[1];
            } else {
                interior[interior.length - 1].right = pieces.left[1];
            }
            interior.push({
                anchor: pieces.left[3],
                left: pieces.left[2],
                right: pieces.right[1]
            });
            currentCurve = pieces.right;
            prevParam = extremaParams[k];
        }
        return { startRight: startRight, endLeft: currentCurve[2], interior: interior };
    }

    /**
     * パスの各曲線セグメントの極点にアンカーを追加します（同一オブジェクトを維持したまま再構築）。
     * @param {PathItem} pathItem - 対象のパス。
     * @returns {number} 追加したアンカー数。
     */
    function addExtremePointsToPath(pathItem) {
        var pts = pathItem.pathPoints;
        var n = pts.length;
        if (n < 2) return 0;

        var isClosed = !!pathItem.closed;
        var segCount = isClosed ? n : (n - 1);

        /* 既存アンカーの左右ハンドル（更新用）と、各セグメント直後に挿入する内部点 */
        var leftHandles = [];
        var rightHandles = [];
        var interiorAfter = [];
        for (var i = 0; i < n; i++) {
            leftHandles[i] = [pts[i].leftDirection[0], pts[i].leftDirection[1]];
            rightHandles[i] = [pts[i].rightDirection[0], pts[i].rightDirection[1]];
            interiorAfter[i] = [];
        }

        var added = 0;
        for (var s = 0; s < segCount; s++) {
            var j = (s + 1) % n;
            var extremaParams = bezierExtremaParams(pts[s], pts[j]);
            if (!extremaParams.length) continue;

            var splitResult = splitSegmentAtParams(
                pts[s].anchor, pts[s].rightDirection,
                pts[j].leftDirection, pts[j].anchor, extremaParams
            );
            rightHandles[s] = splitResult.startRight;
            leftHandles[j] = splitResult.endLeft;
            interiorAfter[s] = splitResult.interior;
            added += splitResult.interior.length;
        }

        if (added === 0) return 0;

        /* 新しい順序で点仕様を組み立てる（既存アンカー＋各セグメント直後の内部点） */
        var newPoints = [];
        for (var a = 0; a < n; a++) {
            newPoints.push({
                anchor: [pts[a].anchor[0], pts[a].anchor[1]],
                left: leftHandles[a],
                right: rightHandles[a],
                type: pts[a].pointType
            });
            var insertedPoints = interiorAfter[a];
            for (var b = 0; b < insertedPoints.length; b++) {
                newPoints.push({
                    anchor: insertedPoints[b].anchor,
                    left: insertedPoints[b].left,
                    right: insertedPoints[b].right,
                    type: PointType.SMOOTH
                });
            }
        }

        /* 同一 PathItem を保持したまま、点数を合わせて全点を上書き（アピアランス維持） */
        while (pathItem.pathPoints.length < newPoints.length) {
            pathItem.pathPoints.add();
        }
        for (var c = 0; c < newPoints.length; c++) {
            var pt = pathItem.pathPoints[c];
            pt.anchor = newPoints[c].anchor;
            pt.pointType = newPoints[c].type;
            pt.leftDirection = newPoints[c].left;
            pt.rightDirection = newPoints[c].right;
        }

        return added;
    }

    /**
     * オブジェクトを再帰的にたどり、含まれるパスの極点にアンカーを追加します。
     * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
     * @returns {void}
     */
    function addExtremePointsToItem(item) {
        if (!item || isSkippableItem(item)) return;
        if (item.typename === 'PathItem') {
            try {
                addExtremePointsToPath(item);
            } catch (e) {
                logProcessError("addExtremePointsToPath", e);
            }
        } else if (item.typename === 'CompoundPathItem') {
            for (var i = 0; i < item.pathItems.length; i++) {
                if (!isSkippableItem(item.pathItems[i])) addExtremePointsToItem(item.pathItems[i]);
            }
        } else if (item.typename === 'GroupItem') {
            for (var j = 0; j < item.pageItems.length; j++) {
                if (!isSkippableItem(item.pageItems[j])) addExtremePointsToItem(item.pageItems[j]);
            }
        }
    }

    /**
     * 指定 targets に対して追加される極点アンカー数を予測します（実際には変更しない）。
     * @param {Array<PathItem>} targets - 対象の PathItem 配列。
     * @returns {number} 追加されるアンカー総数。
     */
    function countExtremaForTargets(targets) {
        var total = 0;
        if (!targets || !targets.length) return total;
        for (var i = 0; i < targets.length; i++) {
            var item = targets[i];
            if (!item || isSkippableItem(item)) continue;
            try {
                var pts = item.pathPoints;
                var n = pts.length;
                if (n < 2) continue;
                var isClosed = !!item.closed;
                var segCount = isClosed ? n : (n - 1);
                for (var s = 0; s < segCount; s++) {
                    var j = (s + 1) % n;
                    total += bezierExtremaParams(pts[s], pts[j]).length;
                }
            } catch (e) {
                logProcessError("countExtremaForTargets", e);
            }
        }
        return total;
    }

    // ダイアログ表示時点の選択を確定（予測と実行を同一対象に揃える）
    var selectionAtOpen = getSelectionOrAlert();
    if (!selectionAtOpen) return;
    var targetsAtOpen = getTargetPathItemsFromSelection(selectionAtOpen);
    var hasHolesAtOpen = selectionHasHoles(selectionAtOpen);

    var ui = showDialog(targetsAtOpen, hasHolesAtOpen);
    if (!ui.ok) return;

    // タブindex依存だと ScriptUI 環境差で誤判定しうるため、明示状態で分岐する
    /**
     * 選択をダイアログ表示時点のスナップショットに復元します。
     * @param {Array<PageItem>} selectionSnapshot - 復元する選択オブジェクト配列。
     * @returns {number} 復元できたオブジェクト数。
     */
    function restoreSelection(selectionSnapshot) {
        var restoredCount = 0;
        try {
            if (!hasDocument()) return 0;
            var doc = app.activeDocument;
            doc.selection = null;
            for (var i = 0; i < selectionSnapshot.length; i++) {
                var item = selectionSnapshot[i];
                if (!item) continue;
                try {
                    item.selected = true;
                    restoredCount++;
                } catch (_) {
                    // skip items that can no longer be selected
                }
            }
        } catch (e) {
            logProcessError("restoreSelection", e);
        }
        return restoredCount;
    }

    /**
     * 選択依存メニューコマンド用に、スキップ対象を除いた PathItem だけを選択復元します。
     * @param {Array<PageItem>} selectionSnapshot - 元の選択スナップショット。
     * @returns {number} 復元できた PathItem 数。
     */
    function restoreSelectableSelection(selectionSnapshot) {
        var restoredCount = 0;
        try {
            if (!hasDocument()) return 0;
            var doc = app.activeDocument;
            var selectablePathItems = getTargetPathItemsFromSelection(selectionSnapshot);

            doc.selection = null;
            for (var i = 0; i < selectablePathItems.length; i++) {
                var pathItem = selectablePathItems[i];
                if (!pathItem || isSkippableItem(pathItem)) continue;
                try {
                    pathItem.selected = true;
                    restoredCount++;
                } catch (_) {
                    // skip items that can no longer be selected
                }
            }
        } catch (e) {
            logProcessError("restoreSelectableSelection", e);
        }
        return restoredCount;
    }

    // 実行前に選択を復元
    if (restoreSelection(selectionAtOpen) === 0) {
        return;
    }

    if (ui.activeMode === 'other') {
        // --- その他タブ: 変換・分割処理 ---
        var selectionSnapshot = selectionAtOpen.slice(0);
        if (ui.convertMode === 'add') {
            if (restoreSelectableSelection(selectionSnapshot) > 0) {
                app.executeMenuCommand('Add Anchor Points2');
            }
        } else if (ui.convertMode === 'split') {
            for (var n = 0; n < selectionSnapshot.length; n++) {
                if (!isSkippableItem(selectionSnapshot[n])) {
                    splitItem(selectionSnapshot[n]);
                }
            }
        } else if (ui.convertMode === 'extreme') {
            for (var n = 0; n < selectionSnapshot.length; n++) {
                if (!isSkippableItem(selectionSnapshot[n])) {
                    addExtremePointsToItem(selectionSnapshot[n]);
                }
            }
        } else if (ui.convertMode === 'fillHoles') {
            if (restoreSelectableSelection(selectionSnapshot) > 0) {
                fillHolesOnSelection();
            }
        } else {
            for (var n = 0; n < selectionSnapshot.length; n++) {
                if (!selectionSnapshot[n] || isSkippableItem(selectionSnapshot[n])) continue;
                convertPathItem(selectionSnapshot[n], ui.convertMode);
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
