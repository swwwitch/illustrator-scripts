#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "ExtendLinesEngine"

/*
概要 / Overview
- 選択オブジェクト内のパス（グループ／複合パス含む）から隣接アンカーポイントのペアを取り、補助線として「直線を描画範囲いっぱいに延長した線」を描画します（線：0.1mm / K100）。
- 「直線」チェックONのときは直線セグメントのみ、OFFのときは曲線セグメントのみを対象にします。
- 選択がアクティブアートボードと交差しない場合、選択の外接幅A・高さBを基準に中心に矩形P（幅A×4、高さB×4）を想定し、その矩形内で延長線を描画します。
- オプション「円弧から円」：Bezier曲線セグメントから円を推定して円を作成します。正確な円弧でない場合は「円弧オプション」で、無視／直線（弦）／直線（延長）を選べます。
- オプション「別レイヤーに」ONのときは `_construction` レイヤーを作成し最前面に配置します（同名レイヤーがある場合は作り直します）。

Overview
- From selected paths (including groups/compound paths), the script takes adjacent anchor pairs and draws auxiliary lines by extending the straight line to fill the drawing bounds (stroke: 0.1mm / K=100).
- When the “Straight” checkbox is ON, only straight segments are processed. When OFF, only curved segments are processed.
- If the selection does not intersect the active artboard, a rectangle P centered on the selection bounds is used instead (width A×4, height B×4), and lines are extended within P.
- Option “Create circle from arc”: estimates a circle from Bezier curve segments and creates circles. If a segment is not a perfect arc, the “Arc options” fallback can be set to Ignore / Straight (chord) / Straight (extend).
- When “Separate layer” is ON, a `_construction` layer is created and brought to front (the layer is recreated if it already exists).

作成日 / Created: 2026-02-27
更新日 / Updated: 2026-02-28
*/



var SCRIPT_VERSION = "v1.0";
var SCRIPT_MARKER = "__ExtendLines__";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "補助線の描画", en: "Extend Lines" },
    historyTitle: { ja: "補助線の描画", en: "Extend Lines" },
    groupName: { ja: "補助線", en: "Aux Lines" },


    modeStraightOnly: { ja: "直線", en: "Straight" },
    countFmt: { ja: "（%n%本）", en: " (%n% lines)" },

    panelOptions: { ja: "オプション", en: "Options" },
    panelAuxLines: { ja: "補助線を描画", en: "Draw guide lines" },
    cbGroup: { ja: "グループ化", en: "Group" },
    cbSeparateLayer: { ja: "別レイヤーに", en: "Separate layer" },
    cbGuide: { ja: "ガイド化", en: "Convert to guides" },
    cbDedup: { ja: "線のダブりを削除", en: "Remove duplicates" },
    cbArcToCircle: { ja: "円弧から円", en: "Create circle from arc" },

    panelArcOptions: { ja: "円弧オプション", en: "Arc options" },
    arcOptionsHint: { ja: "完全な円弧以外の場合", en: "If not a perfect circular arc" },
    arcFallbackIgnore: { ja: "無視", en: "Ignore" },
    arcFallbackStraight: { ja: "直線", en: "Straight" },
    arcFallbackExtend: { ja: "直線（延長）", en: "Straight (extend)" },

    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOk: { ja: "OK", en: "OK" },

    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
    alertNoValidPath: { ja: "有効なパスが見つかりません。", en: "No valid paths were found." },
    alertError: { ja: "エラー: ", en: "Error: " }
};

function L(key) {
    var v = LABELS[key];
    if (!v) return key;
    return v[lang] || v.en || v.ja || key;
}

function fmtCount(n) {
    var s = L("countFmt");
    return s.replace(/%n%/g, String(n));
}

/* セッション保持（Illustrator終了で破棄） / Session-only state (forgotten when Illustrator quits) */
var __EXTENDLINES_SESSION__ = (typeof __EXTENDLINES_SESSION__ !== "undefined") ? __EXTENDLINES_SESSION__ : {};

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var doc = app.activeDocument;
    if (!doc) {
        alert(L("alertNoDoc"));
        return;
    }

    // Undo 1回で元に戻せるように、処理を1つの履歴にまとめる
    try {
        if (doc.suspendHistory) {
            doc.suspendHistory(L("historyTitle"), "mainImpl()");
        } else {
            // 古い環境向けフォールバック
            mainImpl();
        }
    } catch (e) {
        // suspendHistory 内の例外もここに来る
        try { alert(L("alertError") + e); } catch (_) { }
    }
}

function mainImpl() {
    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    // 選択されたオブジェクトからパスを抽出（テキストは一時アウトライン化してから抽出）
    var targetPaths = [];
    var tempOutlineRoots = [];

    // まず通常のパスを抽出
    extractPathItems(sel, targetPaths);

    // テキストがあれば一時アウトライン化してパスを追加
    var tmp = outlineTextFromSelection(sel);
    tempOutlineRoots = tmp.outlineRoots;
    if (tempOutlineRoots.length > 0) {
        extractPathItems(tempOutlineRoots, targetPaths);
    }

    if (targetPaths.length === 0) {
        alert(L("alertNoValidPath"));
        // 一時アウトラインがあれば後始末
        cleanupTempOutlines(tempOutlineRoots);
        return;
    }

    // ダイアログの表示（右側に描画本数を表示するため、事前にカウント）
    var countStraight = countStraightSegments(targetPaths);
    var dialogResult = showDialog(countStraight);
    if (dialogResult === null) {
        cleanupTempOutlines(tempOutlineRoots);
        return; // キャンセルされた場合は終了
    }

    var mode = dialogResult.mode;
    var shouldGroup = dialogResult.group;
    var shouldSeparateLayer = dialogResult.separateLayer;
    var shouldGuide = dialogResult.guide;
    var shouldDedup = dialogResult.dedup;
    var shouldArcToCircle = dialogResult.arcToCircle;
    var arcFallbackMode = dialogResult.arcFallback; // "IGNORE" | "STRAIGHT" | "EXTEND"
    var dedupMap = {}; // 線のダブり検出用

    // 描画範囲（通常はアートボード、選択がアートボード外なら選択中心の矩形P）
    // アートボードの境界 [left, top, right, bottom]
    var abRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var abLeft = abRect[0];
    var abTop = abRect[1];
    var abRight = abRect[2];
    var abBottom = abRect[3];

    // デフォルトはアートボード
    var drawLeft = abLeft;
    var drawTop = abTop;
    var drawRight = abRight;
    var drawBottom = abBottom;

    // 選択全体のバウンディング
    var selBounds = getUnionBounds(sel);
    if (selBounds) {
        var sL = selBounds[0], sT = selBounds[1], sR = selBounds[2], sB = selBounds[3];

        // 選択がアートボードと全く交差しない（完全に外）場合のみ、矩形Pを使う
        // ※Illustrator座標は top > bottom（Yが上に行くほど増える）前提
        var intersects = !(sR < abLeft || sL > abRight || sT < abBottom || sB > abTop);
        if (!intersects) {
            var A = sR - sL; // 選択幅
            var B = sT - sB; // 選択高さ

            // 幅/高さが極端に小さい場合の安全策
            if (A < 1) A = 1;
            if (B < 1) B = 1;

            var cx = (sL + sR) / 2;
            var cy = (sT + sB) / 2;

            var pW = A * 4;
            var pH = B * 4;

            drawLeft = cx - pW / 2;
            drawRight = cx + pW / 2;
            drawTop = cy + pH / 2;
            drawBottom = cy - pH / 2;
        }
    }

    // 描画先レイヤーの決定（別レイヤー対応）
    var targetLayer = doc.activeLayer;

    if (shouldSeparateLayer) {
        var baseLayerName = "_construction";

        // 選択が _construction 上なら、元のオブジェクトを消さないため新規レイヤーを作る
        var mustCreateNew = isSelectionOnLayer(sel, baseLayerName);

        if (mustCreateNew) {
            // 既存の _construction をバックアップ名へリネームし、新しい _construction を作る
            var existing = findLayerByName(doc, baseLayerName);
            if (existing) {
                var backupBase = baseLayerName + "_backup";
                var backupName = createUniqueLayerName(doc, backupBase);
                try { existing.name = backupName; } catch (_) { }
            }

            targetLayer = doc.layers.add();
            targetLayer.name = baseLayerName;
        } else {
            // 既存があれば再利用、なければ新規作成
            targetLayer = findLayerByName(doc, baseLayerName);
            if (!targetLayer) {
                targetLayer = doc.layers.add();
                targetLayer.name = baseLayerName;
            }

            // 再実行時は、このスクリプトが生成したものだけクリア
            clearGeneratedItemsInLayer(targetLayer);
        }

        // 最前面へ移動
        try {
            targetLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (_) { }
    }

    // 新しく引く線の格納先（グループ化オプション対応）
    var lineGroup;
    if (shouldGroup) {
        lineGroup = targetLayer.groupItems.add();
        lineGroup.name = SCRIPT_MARKER + "_" + L("groupName");
        try { lineGroup.note = SCRIPT_MARKER; } catch (_) { }
    } else {
        lineGroup = targetLayer;
    }

    // 円弧から円（オプション）
    if (shouldArcToCircle) {
        for (var c = 0; c < targetPaths.length; c++) {
            try {
                createCirclesFromArcPath(
                    targetPaths[c],
                    lineGroup,
                    shouldGuide,
                    arcFallbackMode,
                    drawLeft,
                    drawTop,
                    drawRight,
                    drawBottom,
                    shouldDedup,
                    dedupMap
                );
            } catch (_) { }
        }
    }

    try {
        for (var p = 0; p < targetPaths.length; p++) {
            var pathItem = targetPaths[p];
            var pts = pathItem.pathPoints;

            if (!pts || pts.length < 2) continue;

            // 隣接するアンカーポイントのペアを生成
            var pairs = [];
            for (var i = 0; i < pts.length - 1; i++) {
                pairs.push([i, i + 1]);
            }

            // パスが閉じている場合は、最後の点と最初の点をつなぐ
            if (pathItem.closed && pts.length >= 3) {
                pairs.push([pts.length - 1, 0]);
            }

            for (var j = 0; j < pairs.length; j++) {
                var pt1 = pts[pairs[j][0]];
                var pt2 = pts[pairs[j][1]];

                // セグメント種別判定
                var segIsStraight = isStraightSegment(pt1, pt2);

                // モードによる対象セグメントの絞り込み
                // STRAIGHT: 直線セグメントのみ
                // CURVE: 曲線セグメントのみ
                if (mode === "STRAIGHT") {
                    if (!segIsStraight) {
                        continue;
                    }
                } else if (mode === "CURVE") {
                    if (segIsStraight) {
                        continue;
                    }
                }

                // 「円弧から円」ON のときは、曲線セグメントの処理は createCirclesFromArcPath 側に委譲する
                // （ここでアンカー同士を結ぶ直線を描くと、円弧が直線になってしまう）
                if (shouldArcToCircle && !segIsStraight) {
                    continue;
                }

                var p1 = pt1.anchor;
                var p2 = pt2.anchor;

                // 2点が全く同じ座標にある場合（ゴミパスなど）はスキップ
                if (Math.abs(p1[0] - p2[0]) < 0.001 && Math.abs(p1[1] - p2[1]) < 0.001) {
                    continue;
                }

                drawLineAcrossArtboard(lineGroup, p1, p2, drawLeft, drawTop, drawRight, drawBottom, shouldGuide, shouldDedup, dedupMap);
            }
        }
    } finally {
        // 一時アウトラインの後始末（失敗時も確実に削除）
        cleanupTempOutlines(tempOutlineRoots);
    }
}

// 選択オブジェクト全体の外接バウンディング（geometricBounds）を取得
// 戻り値: [left, top, right, bottom]
function getUnionBounds(items) {
    var b = null;
    for (var i = 0; i < items.length; i++) {
        try {
            if (!items[i] || !items[i].geometricBounds) continue;
            var gb = items[i].geometricBounds; // [L, T, R, B]
            if (!b) {
                b = [gb[0], gb[1], gb[2], gb[3]];
            } else {
                if (gb[0] < b[0]) b[0] = gb[0];
                if (gb[1] > b[1]) b[1] = gb[1];
                if (gb[2] > b[2]) b[2] = gb[2];
                if (gb[3] < b[3]) b[3] = gb[3];
            }
        } catch (_) { }
    }
    return b;
}

/* ダイアログを表示して設定を取得 / Show dialog and get settings */
function showDialog(countStraight) {
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = ["left", "top"];
    win.margins = 15;

    // Restore last dialog position (session-only)
    try {
        if (__EXTENDLINES_SESSION__.dlgLoc && __EXTENDLINES_SESSION__.dlgLoc.length === 2) {
            win.location = __EXTENDLINES_SESSION__.dlgLoc;
        }
    } catch (_) { }

    // win.add("statictext", undefined, "隣り合うアンカーポイントの扱い：");

    // 2カラムレイアウト
    var cols = win.add("group");
    cols.orientation = "row";
    cols.alignChildren = ["fill", "top"];
    cols.alignment = ["fill", "top"];
    cols.spacing = 10;

    // 補助線を描画 panel
    var pnlExtend = cols.add("panel", undefined, L("panelAuxLines"));
    pnlExtend.orientation = "column";
    pnlExtend.alignChildren = ["left", "top"];
    pnlExtend.margins = [15, 20, 15, 10];
    pnlExtend.alignment = ["fill", "top"]; // 左カラム

    var rbGroup = pnlExtend.add("group");
    rbGroup.orientation = "column";
    rbGroup.alignChildren = ["left", "top"];
    rbGroup.margins = [0, 0, 0, 0];

    // チェックボックス（直線のときだけ）
    var rowStraight = rbGroup.add("group");
    rowStraight.orientation = "row";
    rowStraight.alignChildren = ["left", "center"];
    rowStraight.spacing = 0;

    var cbStraightOnly = rowStraight.add("checkbox", undefined, L("modeStraightOnly"));
    var stStraight = rowStraight.add("statictext", undefined, fmtCount(countStraight));

    // 本数が 0 の場合はディム
    if (countStraight === 0) {
        cbStraightOnly.enabled = false;
        stStraight.enabled = false;
    }

    // デフォルトはON（直線のときだけ）
    cbStraightOnly.value = (countStraight > 0);

    // 追加オプション（補助線を描画 panel 内）
    var cbArcToCircle = rbGroup.add("checkbox", undefined, L("cbArcToCircle"));
    cbArcToCircle.value = true; // デフォルトON

    // 円弧オプション（補助線を描画 panel 内）
    var pnlArcOpt = rbGroup.add("panel", undefined, L("panelArcOptions"));
    pnlArcOpt.orientation = "column";
    pnlArcOpt.alignChildren = ["left", "top"];
    pnlArcOpt.margins = [15, 20, 15, 10];
    pnlArcOpt.helpTip = L("arcOptionsHint");

    var grpArcRb = pnlArcOpt.add("group");
    grpArcRb.orientation = "column";
    grpArcRb.alignChildren = ["left", "top"];

    var rbArcIgnore = grpArcRb.add("radiobutton", undefined, L("arcFallbackIgnore"));
    var rbArcStraight = grpArcRb.add("radiobutton", undefined, L("arcFallbackStraight"));
    var rbArcExtend = grpArcRb.add("radiobutton", undefined, L("arcFallbackExtend"));
    rbArcIgnore.value = true; // デフォルト：無視

    // 念のため排他を強制（環境差で同時ONになる事故を防ぐ）
    function _setArcFallback(which) {
        rbArcIgnore.value = (which === "IGNORE");
        rbArcStraight.value = (which === "STRAIGHT");
        rbArcExtend.value = (which === "EXTEND");
    }
    rbArcIgnore.onClick = function () { _setArcFallback("IGNORE"); };
    rbArcStraight.onClick = function () { _setArcFallback("STRAIGHT"); };
    rbArcExtend.onClick = function () { _setArcFallback("EXTEND"); };

    // cbArcToCircle がOFFのときはパネルをディム
    pnlArcOpt.enabled = cbArcToCircle.value;
    cbArcToCircle.onClick = function () {
        pnlArcOpt.enabled = cbArcToCircle.value;
    };

    // オプションパネル
    var optPanel = cols.add("panel", undefined, L("panelOptions"));
    optPanel.orientation = "column";
    optPanel.alignChildren = ["left", "top"];
    optPanel.margins = [15, 20, 15, 10];
    optPanel.alignment = ["fill", "top"]; // ダイアログ左右いっぱいに

    var cbGroup = optPanel.add("checkbox", undefined, L("cbGroup"));
    cbGroup.value = true; // デフォルトON

    var cbSeparateLayer = optPanel.add("checkbox", undefined, L("cbSeparateLayer"));
    cbSeparateLayer.value = true; // デフォルトON

    var cbGuide = optPanel.add("checkbox", undefined, L("cbGuide"));
    cbGuide.value = false; // デフォルトOFF

    var cbDedup = optPanel.add("checkbox", undefined, L("cbDedup"));
    cbDedup.value = true; // デフォルトON

    // 2カラムの下に余白
    win.add("panel", undefined, undefined);

    var btnGroup = win.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignChildren = ["fill", "center"];
    btnGroup.alignment = ["fill", "top"];

    // 左側にスペーサーを置いて、ボタンを右寄せ（OK を右）
    var spacer = btnGroup.add("group");
    spacer.alignment = ["fill", "center"];
    spacer.minimumSize.width = 0;

    var btnCancel = btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOk = btnGroup.add("button", undefined, L("btnOk"), { name: "ok" });

    var result = null;
    var dialogReturn = win.show();

    // Save dialog position (session-only)
    try {
        __EXTENDLINES_SESSION__.dlgLoc = [win.location[0], win.location[1]];
    } catch (_) { }

    if (dialogReturn === 1) {
        result = {
            mode: cbStraightOnly.value ? "STRAIGHT" : "CURVE",
            group: cbGroup.value,
            separateLayer: cbSeparateLayer.value,
            guide: cbGuide.value,
            dedup: cbDedup.value,
            arcToCircle: cbArcToCircle.value,
            arcFallback: rbArcIgnore.value ? "IGNORE" : (rbArcExtend.value ? "EXTEND" : "STRAIGHT")
        };
    }

    return result;
}

// 選択内のテキストを一時的に複製してアウトライン化し、アウトライン（グループ等）を返す
// 元のテキストは変更しない。テキストが無ければ空配列。
function outlineTextFromSelection(items) {
    var outlineRoots = [];

    function walk(item) {
        if (!item) return;

        if (item.typename === "TextFrame") {
            try {
                // 同一レイヤー末尾に複製（元は触らない）
                var parentContainer = item.layer;
                var dup = item.duplicate(parentContainer, ElementPlacement.PLACEATEND);

                // 複製だけアウトライン化
                var outlined = dup.createOutline();
                // createOutline 後、複製テキストが残る場合があるので削除
                try { dup.remove(); } catch (_) { }

                if (outlined) {
                    outlineRoots.push(outlined);
                }
            } catch (e) {
                // 失敗しても全体は止めない
            }
            return;
        }

        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        } else if (item.typename === "CompoundPathItem") {
            for (var j = 0; j < item.pathItems.length; j++) {
                walk(item.pathItems[j]);
            }
        }
    }

    for (var i = 0; i < items.length; i++) {
        walk(items[i]);
    }

    return { outlineRoots: outlineRoots };
}

// 一時アウトライン（グループ等）を削除
function cleanupTempOutlines(outlineRoots) {
    if (!outlineRoots || outlineRoots.length === 0) return;
    for (var i = outlineRoots.length - 1; i >= 0; i--) {
        try {
            if (outlineRoots[i] && outlineRoots[i].remove) outlineRoots[i].remove();
        } catch (_) { }
    }
}

// 選択が特定レイヤー上かどうか（全選択がそのレイヤーに属する場合のみ true）
function isSelectionOnLayer(items, layerName) {
    try {
        if (!items || items.length === 0) return false;
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) return false;
            var lyr = null;
            try { lyr = it.layer; } catch (_) { lyr = null; }
            if (!lyr || lyr.name !== layerName) return false;
        }
        return true;
    } catch (_) { }
    return false;
}

function findLayerByName(doc, name) {
    try { return doc.layers.getByName(name); } catch (_) { }
    return null;
}

function createUniqueLayerName(doc, baseName) {
    var name = baseName;
    var idx = 2;
    while (findLayerByName(doc, name)) {
        name = baseName + "_" + idx;
        idx++;
    }
    return name;
}

// このスクリプトが生成したオブジェクトだけを削除（ユーザーの既存オブジェクトは残す）
function clearGeneratedItemsInLayer(layer) {
    if (!layer) return;
    try {
        // groupItems を後ろから走査
        for (var i = layer.groupItems.length - 1; i >= 0; i--) {
            var g = layer.groupItems[i];
            if (!g) continue;
            var n = "";
            try { n = g.note || ""; } catch (_) { }
            if (n === SCRIPT_MARKER) {
                try { g.remove(); } catch (_) { }
                continue;
            }
            var nm = "";
            try { nm = g.name || ""; } catch (_) { }
            if (nm.indexOf(SCRIPT_MARKER) === 0) {
                try { g.remove(); } catch (_) { }
            }
        }

        // 直に layer に追加している場合もあるので pathItems も後ろから走査
        for (var j = layer.pathItems.length - 1; j >= 0; j--) {
            var p = layer.pathItems[j];
            if (!p) continue;
            var pn = "";
            try { pn = p.note || ""; } catch (_) { }
            if (pn === SCRIPT_MARKER) {
                try { p.remove(); } catch (_) { }
            }
        }
    } catch (_) { }
}

/* 描画される直線線分（本数）をカウント / Count drawable straight segments */
function countStraightSegments(targetPaths) {
    var straightCount = 0;
    var epsilon = 0.001;

    for (var p = 0; p < targetPaths.length; p++) {
        var pathItem = targetPaths[p];
        var pts = pathItem.pathPoints;
        if (!pts || pts.length < 2) continue;

        // 隣接ペア
        var pairs = [];
        for (var i = 0; i < pts.length - 1; i++) {
            pairs.push([i, i + 1]);
        }
        if (pathItem.closed && pts.length >= 3) {
            pairs.push([pts.length - 1, 0]);
        }

        for (var k = 0; k < pairs.length; k++) {
            var pt1 = pts[pairs[k][0]];
            var pt2 = pts[pairs[k][1]];
            if (!pt1 || !pt2) continue;

            var p1 = pt1.anchor;
            var p2 = pt2.anchor;

            // 同一点は除外（描画側と合わせる）
            if (Math.abs(p1[0] - p2[0]) < epsilon && Math.abs(p1[1] - p2[1]) < epsilon) {
                continue;
            }

            if (isStraightSegment(pt1, pt2)) {
                straightCount++;
            }
        }
    }

    return straightCount;
}

// 2つのアンカーポイント間が直線（ハンドルが出ていない）かどうかを判定する関数
function isStraightSegment(pt1, pt2) {
    var epsilon = 0.001;

    // pt1の出力ハンドル(rightDirection)がアンカーと同じ位置か
    var rDx = Math.abs(pt1.rightDirection[0] - pt1.anchor[0]);
    var rDy = Math.abs(pt1.rightDirection[1] - pt1.anchor[1]);

    // pt2の入力ハンドル(leftDirection)がアンカーと同じ位置か
    var lDx = Math.abs(pt2.leftDirection[0] - pt2.anchor[0]);
    var lDy = Math.abs(pt2.leftDirection[1] - pt2.anchor[1]);

    return (rDx < epsilon && rDy < epsilon && lDx < epsilon && lDy < epsilon);
}

// グループや複合パスの中から再帰的にパスアイテムを抽出する関数
function extractPathItems(items, targetArray) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "PathItem") {
            targetArray.push(item);
        } else if (item.typename === "CompoundPathItem") {
            extractPathItems(item.pathItems, targetArray);
        } else if (item.typename === "GroupItem") {
            extractPathItems(item.pageItems, targetArray);
        }
    }
}

// 2点から線分の重複判定キーを作る（向きは無視）
function makeLineKey(a, b) {
    // ScriptUI/Illustratorの誤差吸収のため、少し丸める
    var ka = pointKey(a);
    var kb = pointKey(b);
    return (ka < kb) ? (ka + "|" + kb) : (kb + "|" + ka);
}

function pointKey(p) {
    // 0.001pt 単位で丸め（十分小さいが、浮動小数の揺れは抑える）
    return (Number(p[0]).toFixed(3) + "," + Number(p[1]).toFixed(3));
}

// 円弧（Bezier曲線セグメント）から円を推定して作成する
// 1セグメント（2アンカー）単位で円を作成する（複数セグメントがあれば複数円）
function getCurvedAdjacentSegments(pathItem) {
    var pts = pathItem.pathPoints;
    var n = pts.length;
    var tol = 1e-6;
    var closed = pathItem.closed;
    var segs = [];

    for (var i = 0; i < n; i++) {
        if (!closed && i === n - 1) break;

        var a = pts[i];
        var b = pts[(i + 1) % n];

        var aRight = a.rightDirection;
        var aAnchor = a.anchor;
        var bLeft = b.leftDirection;
        var bAnchor = b.anchor;

        var aRightDiff = Math.abs(aRight[0] - aAnchor[0]) > tol || Math.abs(aRight[1] - aAnchor[1]) > tol;
        var bLeftDiff = Math.abs(bLeft[0] - bAnchor[0]) > tol || Math.abs(bLeft[1] - bAnchor[1]) > tol;

        if (aRightDiff || bLeftDiff) {
            segs.push({ a: a, b: b });
        }
    }

    return segs;
}

function getCircleCenterFrom2AnchorArc(p1, h1, p3, h2) {
    // Tangent vectors
    var t1 = [h1[0] - p1[0], h1[1] - p1[1]];
    var t2 = [p3[0] - h2[0], p3[1] - h2[1]]; // end tangent direction

    // Normals (radius directions)
    var n1 = rot90(t1);
    var n2 = rot90(t2);

    return lineIntersectionPointDir(p1, n1, p3, n2);
}

function rot90(v) {
    return [-v[1], v[0]];
}

function lineIntersectionPointDir(p, v, q, w) {
    var denom = v[0] * w[1] - v[1] * w[0];
    if (Math.abs(denom) < 1e-9) return null;

    var dx = q[0] - p[0];
    var dy = q[1] - p[1];

    var t = (dx * w[1] - dy * w[0]) / denom;
    return [p[0] + v[0] * t, p[1] + v[1] * t];
}

// 補助線（延長線）と同じスタイルを適用
function applyAuxStyle(pathItem, shouldGuide) {
    try { pathItem.note = SCRIPT_MARKER; } catch (_) { }
    pathItem.filled = false;
    if (shouldGuide) {
        pathItem.stroked = false;
        try { pathItem.guides = true; } catch (_) { }
    } else {
        pathItem.stroked = true;

        var k100 = new CMYKColor();
        k100.cyan = 0;
        k100.magenta = 0;
        k100.yellow = 0;
        k100.black = 100;
        pathItem.strokeColor = k100;

        pathItem.strokeWidth = (0.1 * 72.0) / 25.4;
    }
}

// --- Circle/Arc helpers ---

function dot2(a, b) {
    return a[0] * b[0] + a[1] * b[1];
}

function len2(v) {
    return Math.sqrt(v[0] * v[0] + v[1] * v[1]);
}

function sub2(a, b) {
    return [a[0] - b[0], a[1] - b[1]];
}

// Cubic Bezier point at parameter t (0..1)
function cubicBezierPoint(p0, p1, p2, p3, t) {
    var u = 1 - t;
    var uu = u * u;
    var uuu = uu * u;
    var tt = t * t;
    var ttt = tt * t;

    return [
        uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
        uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
    ];
}

// Heuristic validation: endpoints + midpoint lie on same circle and endpoint tangents are perpendicular to radius
function isApproxCircularArc(p1, h1, h2, p3, center, radius) {
    if (!center || !(radius > 0)) return false;

    // Relative tolerance (1% of radius) with a small absolute floor (0.2pt)
    var tol = Math.max(0.2, radius * 0.01);

    // Distances at endpoints
    var d1 = len2(sub2(p1, center));
    var d3 = len2(sub2(p3, center));
    if (Math.abs(d1 - radius) > tol) return false;
    if (Math.abs(d3 - radius) > tol) return false;

    // Midpoint on Bezier
    var mid = cubicBezierPoint(p1, h1, h2, p3, 0.5);
    var dm = len2(sub2(mid, center));
    if (Math.abs(dm - radius) > tol) return false;

    // Tangent orthogonality at endpoints (radius · tangent ≈ 0)
    var t1 = sub2(h1, p1);          // start tangent direction
    var t2 = sub2(p3, h2);          // end tangent direction
    var r1 = sub2(p1, center);
    var r3 = sub2(p3, center);

    // If tangents are too small, treat as invalid (likely not a proper arc)
    if (len2(t1) < 1e-6 || len2(t2) < 1e-6) return false;

    if (Math.abs(dot2(r1, t1)) > tol * len2(t1)) return false;
    if (Math.abs(dot2(r3, t2)) > tol * len2(t2)) return false;

    return true;
}

function createChordLine(targetContainer, p1, p3, shouldGuide) {
    try {
        var ln = targetContainer.pathItems.add();
        ln.setEntirePath([p1, p3]);
        ln.closed = false;
        applyAuxStyle(ln, shouldGuide);
        return ln;
    } catch (_) { }
    return null;
}

function createCirclesFromArcPath(arcPath, targetContainer, shouldGuide, arcFallbackMode, drawLeft, drawTop, drawRight, drawBottom, shouldDedup, dedupMap) {
    var circles = [];
    try {
        var segs = getCurvedAdjacentSegments(arcPath);
        if (!segs || segs.length === 0) return circles;

        for (var i = 0; i < segs.length; i++) {
            var a = segs[i].a;
            var b = segs[i].b;

            var p1 = a.anchor;
            var p3 = b.anchor;
            var h1 = a.rightDirection;
            var h2 = b.leftDirection;

            var center = getCircleCenterFrom2AnchorArc(p1, h1, p3, h2);
            if (!center) continue;

            var dx = p1[0] - center[0];
            var dy = p1[1] - center[1];
            var radius = Math.sqrt(dx * dx + dy * dy);
            if (!(radius > 0)) continue;

            // 正確な円弧の一部でない場合の扱い
            if (!isApproxCircularArc(p1, h1, h2, p3, center, radius)) {
                if (arcFallbackMode === "STRAIGHT") {
                    // 直線：アンカー同士を結ぶ弦
                    createChordLine(targetContainer, p1, p3, shouldGuide);
                } else if (arcFallbackMode === "EXTEND") {
                    // 直線（延長）：アンカー同士を結ぶ直線を描画範囲いっぱいに延長
                    // drawLineAcrossArtboard は dedupMap を使って重複線を抑制できる
                    drawLineAcrossArtboard(
                        targetContainer,
                        p1,
                        p3,
                        drawLeft,
                        drawTop,
                        drawRight,
                        drawBottom,
                        shouldGuide,
                        shouldDedup,
                        dedupMap
                    );
                }
                // IGNORE は何もしない
                continue;
            }

            // Create ellipse (circle)
            var circle = targetContainer.pathItems.ellipse(
                center[1] + radius,
                center[0] - radius,
                radius * 2,
                radius * 2
            );
            applyAuxStyle(circle, shouldGuide);
            circles.push(circle);
        }
    } catch (_) { }

    return circles;
}

// 延長線を計算して描画する関数
function drawLineAcrossArtboard(targetGroup, p1, p2, left, top, right, bottom, shouldGuide, shouldDedup, dedupMap) {
    var x1 = p1[0], y1 = p1[1];
    var x2 = p2[0], y2 = p2[1];

    var intersections = [];
    var epsilon = 0.001;

    // 垂直線の場合
    if (Math.abs(x1 - x2) < epsilon) {
        intersections.push([x1, top]);
        intersections.push([x1, bottom]);
    }
    // 水平線の場合
    else if (Math.abs(y1 - y2) < epsilon) {
        intersections.push([left, y1]);
        intersections.push([right, y1]);
    }
    // それ以外（傾きがある場合）
    else {
        var m = (y2 - y1) / (x2 - x1);
        var b = y1 - m * x1;

        // x = left のときの y
        var yAtLeft = m * left + b;
        if (yAtLeft <= top + epsilon && yAtLeft >= bottom - epsilon) intersections.push([left, yAtLeft]);

        // x = right のときの y
        var yAtRight = m * right + b;
        if (yAtRight <= top + epsilon && yAtRight >= bottom - epsilon) intersections.push([right, yAtRight]);

        // y = top のときの x
        var xAtTop = (top - b) / m;
        if (xAtTop >= left - epsilon && xAtTop <= right + epsilon) intersections.push([xAtTop, top]);

        // y = bottom のときの x
        var xAtBottom = (bottom - b) / m;
        if (xAtBottom >= left - epsilon && xAtBottom <= right + epsilon) intersections.push([xAtBottom, bottom]);
    }

    // 交点が2つ以上見つかった場合、直線を引く
    if (intersections.length >= 2) {
        // 重複する交点を削除
        var uniquePoints = [intersections[0]];
        for (var j = 1; j < intersections.length; j++) {
            var dx = intersections[j][0] - uniquePoints[0][0];
            var dy = intersections[j][1] - uniquePoints[0][1];
            if (Math.sqrt(dx * dx + dy * dy) > epsilon) {
                uniquePoints.push(intersections[j]);
                break;
            }
        }

        if (uniquePoints.length === 2) {
            // 線のダブりを削除（同じ2点で構成される延長線は1本だけ描く）
            if (shouldDedup && dedupMap) {
                var key = makeLineKey(uniquePoints[0], uniquePoints[1]);
                if (dedupMap[key]) {
                    return;
                }
                dedupMap[key] = true;
            }
            var newLine = targetGroup.pathItems.add();
            newLine.setEntirePath(uniquePoints);
            newLine.closed = false;
            applyAuxStyle(newLine, shouldGuide);

        }
    }
}

main();
