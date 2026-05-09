#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要 / Overview

- 選択内容に応じて 2 つのモードを自動判定する Illustrator 用スクリプト
  1. 長方形 → 中心線（縦線または横線）に変換
  2. 4 本の水平／垂直線が「#」状に重なっている場合 → 長方形に変換
- いずれのモードも元オブジェクトは削除され、新しいオブジェクトが選択される

- An Illustrator script with two auto-detected modes:
  1. Rectangle → center line (horizontal or vertical)
  2. Four horizontal/vertical lines forming a "#" pattern → rectangle
- In either mode, the source objects are deleted and the new objects become selected.

### 主な機能 / Main Features

- 選択内容からモードを自動判定（両モードを同時に実行することも可能）
- 縦長／横長の長方形に対応し、適切な方向の中心線を描画
- 回転補正（任意でON/OFF）で角度ズレを修正
- 4 本の H/V 線が交差して「#」を形成する組み合わせを検出して長方形化（最も太い線の属性を継承）
- ライブプレビュー対応
- 除外条件（正方形に近い形状、短辺×1.5＞長辺）あり
- 日本語／英語インターフェース対応

- Auto-detects mode from the current selection (both can run simultaneously)
- Supports vertical / horizontal rectangles, draws appropriate center line
- Optional rotation correction to fix angle misalignments
- Detects four H/V lines forming a "#" pattern and converts them into a rectangle (inheriting the thickest line's attributes)
- Live preview
- Exclusion conditions (near-square, short × 1.5 > long)
- Japanese / English UI support

### 処理の流れ / Process Flow

1. 選択から「閉じた4点パスの長方形」と「2点の H/V 直線」を抽出
2. 長方形は中央に線を描画し、元の長方形を削除
3. 2H+2V の直線が交差して # を形成する場合は、外周長方形を作成し元の線を削除

### 更新履歴 / Update History

- v1.0.0 (20250612) : 初版作成 / Initial version
- v1.0.1 (20250615) : 除外条件追加、回転補正改善 / Added exclusions, improved rotation
- v1.0.2 (20260426) : ダイアログ追加、ローカライズ整備 / Added dialog, full localization
- v1.1.0 (20260426) : プレビュー機能追加 / Added live preview
- v1.2.0 (20260426) : 「#」状の4本線→長方形モードを統合 / Integrated "#" pattern → rectangle mode

*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitleRect:  { ja: "中心線を描画",                       en: "Draw Center Line" },
    dialogTitleHash:  { ja: "長方形を作成",                       en: "Build Rectangle" },
    dialogTitleBoth:  { ja: "長方形⇄中心線",                      en: "Rectangle ⇄ Center Line" },
    chkRotate:        { ja: "回転を補正",                         en: "Correct rotation" },
    chkPreview:       { ja: "プレビュー",                         en: "Preview" },
    btnOutlineOn:     { ja: "アウトライン表示",                   en: "Outline View" },
    btnOutlineOff:    { ja: "プレビュー表示",                     en: "Preview View" },
    btnOk:            { ja: "OK",                                 en: "OK" },
    btnCancel:        { ja: "キャンセル",                         en: "Cancel" },
    statusRect:       { ja: "長方形 %d 個を検出",                 en: "%d rectangle(s) detected" },
    statusHash:       { ja: "「#」パターン %d 組を検出",           en: "%d \"#\" pattern(s) detected" },
    alertNoTarget:    { ja: "長方形、または「#」状に重なる H/V 線 4 本を選択してください。", en: "Please select rectangles, or 4 H/V lines forming a \"#\" pattern." },
    alertNoSelection: { ja: "長方形を1つ以上選択してください。",   en: "Please select at least one rectangle." },
    alertError:       { ja: "エラーが発生しました",               en: "An error occurred" }
};

/* ラベル取得 / Get label */
function L(key) {
    return LABELS[key][lang];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (lang === 'ja' ? '：' : ':');
}

/* %d を整数で置換した簡易フォーマッタ / Tiny formatter that replaces %d with an integer */
function fmt(template, value) {
    return template.replace(/%d/, String(value));
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* 検出モードからダイアログタイトルを決定 / Pick dialog title from detected modes */
function getDialogTitleKey(hasRects, hasHash) {
    if (hasRects && hasHash) return 'dialogTitleBoth';
    if (hasHash) return 'dialogTitleHash';
    return 'dialogTitleRect';
}

/* オプション設定ダイアログ。プレビュー ON でライブ実行、OK で確定、キャンセルで巻き戻し
   Options dialog: live preview while open, OK commits, Cancel reverts */
function showOptionDialog(rects, hashGroups) {
    var hasRects = rects.length > 0;
    var hasHash  = hashGroups.length > 0;

    var dlg = new Window('dialog', L(getDialogTitleKey(hasRects, hasHash)) + ' ' + SCRIPT_VERSION);
    dlg.alignChildren = "fill";
    dlg.margins = 16;
    dlg.spacing = 12;

    /* 検出状況の表示 / Detection summary */
    var statusGroup = dlg.add("group");
    statusGroup.orientation = "column";
    statusGroup.alignChildren = ["left", "center"];
    if (hasRects) statusGroup.add("statictext", undefined, fmt(L('statusRect'), rects.length));
    if (hasHash)  statusGroup.add("statictext", undefined, fmt(L('statusHash'), hashGroups.length));

    /* 回転補正チェックボックス（長方形検出時のみ）/ Rotation correction (only when rectangles are detected) */
    var cbRotate = null;
    if (hasRects) {
        cbRotate = dlg.add("checkbox", undefined, L('chkRotate'));
        cbRotate.value = true;
    }

    /* プレビューチェックボックス / Preview checkbox */
    var cbPreview = dlg.add("checkbox", undefined, L('chkPreview'));
    cbPreview.value = false;

    /* ボタン行（左：表示切替、右：OK/キャンセル）/ Button row (left: view toggle, right: OK/Cancel) */
    var btnRow = dlg.add("group");
    btnRow.alignment = "fill";
    btnRow.alignChildren = ["fill", "center"];

    var grpLeft = btnRow.add("group");
    grpLeft.alignment = ["left", "center"];

    var isOutlineMode = false;
    var btnView = grpLeft.add("button", undefined, L('btnOutlineOn'));
    btnView.onClick = function () {
        try {
            app.executeMenuCommand('preview');
            isOutlineMode = !isOutlineMode;
            btnView.text = isOutlineMode ? L('btnOutlineOff') : L('btnOutlineOn');
        } catch (e) {}
    };

    var grpRight = btnRow.add("group");
    grpRight.alignment = ["right", "center"];
    var btnCancel = grpRight.add("button", undefined, L('btnCancel'), { name: "cancel" });
    var btnOk     = grpRight.add("button", undefined, L('btnOk'),     { name: "ok" });

    /* プレビュー状態 / Preview state */
    var lineEntries = []; /* 長方形→中心線 / rect → line */
    var rectEntries = []; /* 「#」→ 長方形 / hash → rect */

    function clearAllPreview() {
        clearPreview(lineEntries);
        clearHashPreview(rectEntries);
        lineEntries = [];
        rectEntries = [];
    }

    function refreshPreview() {
        clearAllPreview();
        if (cbPreview.value) {
            if (hasRects) lineEntries = buildPreview(rects, cbRotate ? cbRotate.value : false);
            if (hasHash)  rectEntries = buildHashPreview(hashGroups);
        }
        app.redraw();
    }

    cbPreview.onClick = refreshPreview;
    if (cbRotate) {
        cbRotate.onClick = function () {
            if (cbPreview.value) refreshPreview();
        };
    }

    btnOk.onClick = function () {
        var newItems = [];

        /* 長方形 → 中心線 / Rectangle → center line */
        if (hasRects) {
            if (cbPreview.value && lineEntries.length > 0) {
                var lines = commitFromPreview(lineEntries);
                for (var i = 0; i < lines.length; i++) newItems.push(lines[i]);
            } else {
                clearPreview(lineEntries);
                for (var i = 0; i < rects.length; i++) {
                    var line = drawCenterLineFromRect(rects[i], cbRotate ? cbRotate.value : false);
                    if (line) newItems.push(line);
                }
            }
        }

        /* 「#」→ 長方形 / Hash pattern → rectangle */
        if (hasHash) {
            if (cbPreview.value && rectEntries.length > 0) {
                var newRects = commitHashFromPreview(rectEntries);
                for (var i = 0; i < newRects.length; i++) newItems.push(newRects[i]);
            } else {
                clearHashPreview(rectEntries);
                for (var i = 0; i < hashGroups.length; i++) {
                    var rect = buildRectFromGroup(hashGroups[i]);
                    /* プレビュー無し時は元の線をここで削除 / Without preview, remove source lines now */
                    for (var j = 0; j < hashGroups[i].length; j++) {
                        try { hashGroups[i][j].remove(); } catch (_) {}
                    }
                    if (rect) newItems.push(rect);
                }
            }
        }

        lineEntries = [];
        rectEntries = [];
        if (newItems.length > 0) {
            app.activeDocument.selection = newItems;
        }
        dlg.close(1);
    };

    btnCancel.onClick = function () {
        clearAllPreview();
        app.redraw();
        dlg.close(2);
    };

    dlg.show();

    /* X ボタン等で閉じられた場合の安全網 / Safety net for close-box dismissal */
    if (lineEntries.length > 0 || rectEntries.length > 0) {
        clearAllPreview();
        app.redraw();
    }
}

// =========================================
// 中心線描画処理 / Center Line Drawing
// =========================================

/* 回転補正で必要な角度を返す（補正不要なら 0）/ Return rotation correction angle (0 if not applicable) */
function getRotationCorrectionAngle(rect) {
    if (rect.pathPoints.length !== 4) return 0;
    var ptA = rect.pathPoints[0].anchor;
    var ptB = rect.pathPoints[1].anchor;
    var dx = ptB[0] - ptA[0];
    var dy = ptB[1] - ptA[1];
    var angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
    if (angleDeg < 0) angleDeg += 360;
    var normalized = angleDeg % 90;
    if (normalized > 45) normalized = 90 - normalized;
    if (normalized >= 0.5 && normalized <= 10) return angleDeg;
    return 0;
}

/* 中心線を生成（除外判定込み）。removeOriginal=true で元の長方形を削除
   Create center line (with exclusion check). When removeOriginal=true, also delete the source rect */
function processRect(rect, correctRotation, removeOriginal) {
    var rotatedBy = 0;
    if (correctRotation) {
        var angle = getRotationCorrectionAngle(rect);
        if (angle > 0) {
            rect.rotate(-angle);
            rotatedBy = angle;
        }
    }

    var b = rect.geometricBounds;
    var left = b[0], top = b[1], right = b[2], bottom = b[3];
    var w = right - left;
    var h = top - bottom;

    var diffRatio = Math.abs(w - h) / Math.max(w, h);
    var shortSide = Math.min(w, h);
    var longSide = Math.max(w, h);
    if (diffRatio < 0.05 || shortSide * 1.5 > longSide) {
        /* プレビュー時は回転を巻き戻す / Restore rotation when previewing */
        if (!removeOriginal && rotatedBy !== 0) rect.rotate(rotatedBy);
        return null;
    }

    var centerLine = app.activeDocument.pathItems.add();
    centerLine.stroked = true;
    centerLine.filled = false;
    centerLine.strokeColor = rect.fillColor;

    var p1 = centerLine.pathPoints.add();
    var p2 = centerLine.pathPoints.add();

    if (h <= w) {
        var centerY = (top + bottom) / 2;
        p1.anchor = [left, centerY];
        p2.anchor = [right, centerY];
    } else {
        var centerX = (left + right) / 2;
        p1.anchor = [centerX, top];
        p2.anchor = [centerX, bottom];
    }

    centerLine.strokeWidth = (h <= w) ? h : w;

    p1.leftDirection  = p1.anchor;
    p1.rightDirection = p1.anchor;
    p2.leftDirection  = p2.anchor;
    p2.rightDirection = p2.anchor;

    if (removeOriginal) rect.remove();

    return { line: centerLine, rotatedBy: rotatedBy };
}

/* 元の API 互換ラッパー / Wrapper preserving original API */
function drawCenterLineFromRect(rect, correctRotation) {
    var r = processRect(rect, correctRotation, true);
    return r ? r.line : null;
}

// =========================================
// 選択処理とプレビュー管理 / Selection & Preview Management
// =========================================

/* 選択から処理可能な長方形 PathItem を抽出 / Extract processable rectangle PathItems from selection */
function getProcessableRects(selection) {
    var rects = [];
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "PathItem" && item.closed && item.pathPoints.length === 4) {
            rects.push(item);
        } else if (item.typename === "CompoundPathItem" && item.pathItems.length === 1) {
            var sub = item.pathItems[0];
            if (sub.closed && sub.pathPoints.length === 4) rects.push(sub);
        }
    }
    return rects;
}

/* プレビュー線を作成（元の長方形は残す）/ Build preview lines without removing originals */
function buildPreview(rects, correctRotation) {
    var entries = [];
    for (var i = 0; i < rects.length; i++) {
        var r = processRect(rects[i], correctRotation, false);
        if (r) entries.push({ line: r.line, rect: rects[i], rotatedBy: r.rotatedBy });
    }
    return entries;
}

/* プレビューを破棄して元の状態に戻す / Discard preview and restore originals */
function clearPreview(entries) {
    for (var i = entries.length - 1; i >= 0; i--) {
        var e = entries[i];
        try { e.line.remove(); } catch (_) {}
        if (e.rotatedBy !== 0) {
            try { e.rect.rotate(e.rotatedBy); } catch (_) {}
        }
    }
}

/* プレビュー線をそのまま採用し、元の長方形のみ削除 / Keep preview lines, remove only the source rects */
function commitFromPreview(entries) {
    var lines = [];
    for (var i = 0; i < entries.length; i++) {
        try { entries[i].rect.remove(); } catch (_) {}
        lines.push(entries[i].line);
    }
    return lines;
}

// =========================================
// 「#」状の4本線→長方形 / "#" pattern → Rectangle
// =========================================

/* 選択から「2点の水平／垂直 PathItem」を抽出 / Extract 2-point H/V PathItems from selection */
function getStraightLines(selection) {
    var out = [];
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename !== "PathItem" || item.pathPoints.length !== 2) continue;
        var p1 = item.pathPoints[0].anchor;
        var p2 = item.pathPoints[1].anchor;
        var isH = Math.abs(p1[1] - p2[1]) < 0.01;
        var isV = Math.abs(p1[0] - p2[0]) < 0.01;
        if (isH || isV) out.push({ line: item, isH: isH });
    }
    return out;
}

/* 線分が交差するか（端点接触は除外）/ Strict segment intersection (endpoint touch excluded) */
function segmentsIntersect(a, b) {
    var a1 = a.pathPoints[0].anchor, a2 = a.pathPoints[1].anchor;
    var b1 = b.pathPoints[0].anchor, b2 = b.pathPoints[1].anchor;
    function ccw(p1, p2, p3) {
        return (p3[1] - p1[1]) * (p2[0] - p1[0]) > (p2[1] - p1[1]) * (p3[0] - p1[0]);
    }
    return (ccw(a1, b1, b2) !== ccw(a2, b1, b2)) && (ccw(a1, a2, b1) !== ccw(a1, a2, b2));
}

/* 「#」を構成する 2H+2V の組を検出（同じ線は重複利用しない）
   Detect 2H+2V groups forming "#" patterns (no line is reused) */
function findHashGroups(selection) {
    var classified = getStraightLines(selection);
    var hLines = [], vLines = [];
    for (var i = 0; i < classified.length; i++) {
        if (classified[i].isH) hLines.push(classified[i].line);
        else vLines.push(classified[i].line);
    }
    if (hLines.length < 2 || vLines.length < 2) return [];

    var consumed = [];
    function isConsumed(item) {
        for (var k = 0; k < consumed.length; k++) {
            if (consumed[k] === item) return true;
        }
        return false;
    }

    var groups = [];
    for (var i = 0; i < hLines.length - 1; i++) {
        if (isConsumed(hLines[i])) continue;
        for (var j = i + 1; j < hLines.length; j++) {
            if (isConsumed(hLines[i])) break;
            if (isConsumed(hLines[j])) continue;
            for (var m = 0; m < vLines.length - 1; m++) {
                if (isConsumed(hLines[i]) || isConsumed(hLines[j])) break;
                if (isConsumed(vLines[m])) continue;
                for (var n = m + 1; n < vLines.length; n++) {
                    if (isConsumed(vLines[n])) continue;
                    var h1 = hLines[i], h2 = hLines[j], v1 = vLines[m], v2 = vLines[n];
                    if (segmentsIntersect(h1, v1) && segmentsIntersect(h1, v2) &&
                        segmentsIntersect(h2, v1) && segmentsIntersect(h2, v2)) {
                        groups.push([h1, h2, v1, v2]);
                        consumed.push(h1, h2, v1, v2);
                        break;
                    }
                }
            }
        }
    }
    return groups;
}

/* 4本線の中で最も太い線を参照属性として返す / Return the thickest line as attribute reference */
function pickThickestLine(group) {
    var ref = group[0];
    for (var i = 1; i < group.length; i++) {
        if (group[i].strokeWidth > ref.strokeWidth) ref = group[i];
    }
    return ref;
}

/* 4本線の外周から長方形パスを作成（元の線は残す）
   Build a rectangle path from 4 lines (originals are kept; caller decides removal) */
function buildRectFromGroup(group) {
    var hLines = [], vLines = [];
    for (var i = 0; i < group.length; i++) {
        var p1 = group[i].pathPoints[0].anchor;
        var p2 = group[i].pathPoints[1].anchor;
        if (Math.abs(p1[1] - p2[1]) < 0.01) hLines.push(p1[1]);
        else vLines.push(p1[0]);
    }
    var topY    = Math.max(hLines[0], hLines[1]);
    var bottomY = Math.min(hLines[0], hLines[1]);
    var leftX   = Math.min(vLines[0], vLines[1]);
    var rightX  = Math.max(vLines[0], vLines[1]);

    var ref = pickThickestLine(group);
    var rect = app.activeDocument.pathItems.add();
    rect.stroked = true;
    rect.filled = false;
    rect.strokeWidth = ref.strokeWidth;
    rect.strokeColor = ref.strokeColor;

    var pts = [
        [leftX,  topY],
        [rightX, topY],
        [rightX, bottomY],
        [leftX,  bottomY]
    ];
    for (var k = 0; k < 4; k++) {
        var pt = rect.pathPoints.add();
        pt.anchor = pts[k];
        pt.leftDirection = pts[k];
        pt.rightDirection = pts[k];
        pt.pointType = PointType.CORNER;
    }
    rect.closed = true;
    return rect;
}

/* プレビュー長方形を作成（元の線は残す）/ Build preview rectangles without removing source lines */
function buildHashPreview(groups) {
    var entries = [];
    for (var i = 0; i < groups.length; i++) {
        var rect = buildRectFromGroup(groups[i]);
        entries.push({ rect: rect, lines: groups[i] });
    }
    return entries;
}

/* プレビュー長方形を破棄 / Discard preview rectangles */
function clearHashPreview(entries) {
    for (var i = entries.length - 1; i >= 0; i--) {
        try { entries[i].rect.remove(); } catch (_) {}
    }
}

/* プレビュー長方形を採用し、元の線のみ削除 / Keep preview rectangles, remove only source lines */
function commitHashFromPreview(entries) {
    var rects = [];
    for (var i = 0; i < entries.length; i++) {
        var e = entries[i];
        for (var j = 0; j < e.lines.length; j++) {
            try { e.lines[j].remove(); } catch (_) {}
        }
        rects.push(e.rect);
    }
    return rects;
}

// =========================================
// メイン処理 / Main
// =========================================

function main() {
    try {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(L('alertNoTarget'));
            return;
        }

        var selection   = app.activeDocument.selection;
        var rects       = getProcessableRects(selection);
        var hashGroups  = findHashGroups(selection);

        if (rects.length === 0 && hashGroups.length === 0) {
            alert(L('alertNoTarget'));
            return;
        }

        showOptionDialog(rects, hashGroups);
    } catch (e) {
        alert(labelText('alertError') + "\n" + e);
    }
}

main();