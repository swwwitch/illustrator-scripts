#target illustrator

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "ExtendLinesEngine"

/*
概要 / Overview
- 選択オブジェクト内のパス（グループ／複合パス含む）からアンカーポイントを抽出し、隣接ポイント間の直線をアートボード端まで延長した線として描画します（線：0.1mm / K100）。
- モード：
  - ALL：曲線セグメントも含めて全ての隣接ポイントを対象
  - STRAIGHT：直線セグメントのみ（ハンドルが出ている曲線セグメントは除外）

作成日 / Created: 2026-02-27
更新日 / Updated: 2026-02-27
*/


var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "延長線の描画", en: "Extend Lines" },

    modeAll: { ja: "曲線を含む", en: "Include curves" },
    modeStraight: { ja: "直線のときだけ", en: "Straight only" },
    countFmt: { ja: "（%n%本）", en: " (%n% lines)" },

    panelOptions: { ja: "オプション", en: "Options" },
    cbGroup: { ja: "グループ化", en: "Group" },
    cbSeparateLayer: { ja: "別レイヤーに", en: "Separate layer" },
    cbDedup: { ja: "線のダブりを削除", en: "Remove duplicates" },

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
            doc.suspendHistory("Extend Lines", "mainImpl()");
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
    var counts = countDrawableSegments(targetPaths);
    var dialogResult = showDialog(counts.all, counts.straight);
    if (dialogResult === null) {
        cleanupTempOutlines(tempOutlineRoots);
        return; // キャンセルされた場合は終了
    }

    var mode = dialogResult.mode;
    var shouldGroup = dialogResult.group;
    var shouldSeparateLayer = dialogResult.separateLayer;
    var shouldDedup = dialogResult.dedup;
    var dedupMap = {}; // 線のダブり検出用

    // アートボードの境界を取得 [left, top, right, bottom]
    var abRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var abLeft = abRect[0];
    var abTop = abRect[1];
    var abRight = abRect[2];
    var abBottom = abRect[3];

    // 描画先レイヤーの決定（別レイヤー対応）
    var targetLayer = doc.activeLayer;

    if (shouldSeparateLayer) {
        var layerName = "_construction";

        // 既存レイヤーがあれば削除（再実行で増殖させない）
        try {
            var existingLayer = doc.layers.getByName(layerName);
            if (existingLayer) {
                existingLayer.remove();
            }
        } catch (_) { }

        // 新規作成
        targetLayer = doc.layers.add();
        targetLayer.name = layerName;

        // 最背面へ移動
        try {
            targetLayer.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (_) { }
    }

    // 新しく引く線の格納先（グループ化オプション対応）
    var lineGroup;
    if (shouldGroup) {
        lineGroup = targetLayer.groupItems.add();
        lineGroup.name = "Extended Lines";
    } else {
        lineGroup = targetLayer;
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

                // 「直線のときだけ」モードの場合、ハンドルが出ていればスキップ
                if (mode === "STRAIGHT") {
                    if (!isStraightSegment(pt1, pt2)) {
                        continue;
                    }
                }

                var p1 = pt1.anchor;
                var p2 = pt2.anchor;

                // 2点が全く同じ座標にある場合（ゴミパスなど）はスキップ
                if (Math.abs(p1[0] - p2[0]) < 0.001 && Math.abs(p1[1] - p2[1]) < 0.001) {
                    continue;
                }

                drawLineAcrossArtboard(lineGroup, p1, p2, abLeft, abTop, abRight, abBottom, shouldDedup, dedupMap);
            }
        }
    } finally {
        // 一時アウトラインの後始末（失敗時も確実に削除）
        cleanupTempOutlines(tempOutlineRoots);
    }
}

/* ダイアログを表示して設定を取得 / Show dialog and get settings */
function showDialog(countAll, countStraight) {
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

    var rbGroup = win.add("group");
    rbGroup.orientation = "column";
    rbGroup.alignChildren = ["left", "top"];
    rbGroup.margins = [10, 5, 0, 10];

    // ラジオボタン + 右側に本数表示
    var rowAll = rbGroup.add("group");
    rowAll.orientation = "row";
    rowAll.alignChildren = ["left", "center"];
    rowAll.spacing = 0;
    var rbAll = rowAll.add("radiobutton", undefined, L("modeAll"));
    var stAll = rowAll.add("statictext", undefined, fmtCount(countAll));

    var rowStraight = rbGroup.add("group");
    rowStraight.orientation = "row";
    rowStraight.alignChildren = ["left", "center"];
    rowStraight.spacing = 0;
    var rbStraight = rowStraight.add("radiobutton", undefined, L("modeStraight"));
    var stStraight = rowStraight.add("statictext", undefined, fmtCount(countStraight));

    // 本数が 0 の場合はディム
    if (countAll === 0) {
        rbAll.enabled = false;
        stAll.enabled = false;
    }
    if (countStraight === 0) {
        rbStraight.enabled = false;
        stStraight.enabled = false;
    }

    // デフォルトは「直線のときだけ」
    if (countStraight > 0) {
        rbStraight.value = true;
    } else if (countAll > 0) {
        rbAll.value = true;
    }

    // オプションパネル
    var optPanel = win.add("panel", undefined, L("panelOptions"));
    optPanel.orientation = "column";
    optPanel.alignChildren = ["left", "top"];
    optPanel.margins = [15, 20, 15, 10];

    var cbGroup = optPanel.add("checkbox", undefined, L("cbGroup"));
    cbGroup.value = true; // デフォルトON

    var cbSeparateLayer = optPanel.add("checkbox", undefined, L("cbSeparateLayer"));
    cbSeparateLayer.value = true; // デフォルトON

    var cbDedup = optPanel.add("checkbox", undefined, L("cbDedup"));
    cbDedup.value = true; // デフォルトON

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
            mode: rbAll.value ? "ALL" : "STRAIGHT",
            group: cbGroup.value,
            separateLayer: cbSeparateLayer.value,
            dedup: cbDedup.value
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

/* 描画される線分（本数）をモード別にカウント / Count drawable segments by mode */
function countDrawableSegments(targetPaths) {
    var allCount = 0;
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

            allCount++;
            if (isStraightSegment(pt1, pt2)) {
                straightCount++;
            }
        }
    }

    return { all: allCount, straight: straightCount };
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

// 延長線を計算して描画する関数
function drawLineAcrossArtboard(targetGroup, p1, p2, left, top, right, bottom, shouldDedup, dedupMap) {
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
            newLine.filled = false;
            newLine.stroked = true;

            // 線の色：K100（CMYK）
            var k100 = new CMYKColor();
            k100.cyan = 0;
            k100.magenta = 0;
            k100.yellow = 0;
            k100.black = 100;
            newLine.strokeColor = k100;

            // 線幅：0.1mm
            newLine.strokeWidth = (0.1 * 72.0) / 25.4; // mm -> pt
        }
    }
}

main();
