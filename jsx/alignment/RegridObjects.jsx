#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script name:
RegridObjects

### GitHub:
https://github.com/swwwitch/illustrator-scripts

### 概要 / Overview:
- 選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します。
- ダイアログは日本語／英語の自動切り替えに対応（$.locale）。
- ［連動］で左右の値を上下に反映できます。
- ［行列入れ替え］をONにすると、歯抜け（欠け）を許容しつつ行⇄列を転置します。

### 更新日 / Updated:
- 2026-01-25

Illustrator Grid Re-spacing Script (always-on preview, link option)

- Assumes the selected objects are arranged in an "approximately grid-like" layout.
- Repositions them using the horizontal/vertical gaps entered in the dialog.
- The dialog supports automatic Japanese/English switching (based on $.locale).
- The "Link" option mirrors the horizontal value to the vertical value.
- When "Swap Rows/Columns" is ON, it transposes rows/columns while tolerating missing cells.

### Updated:
- 2026-01-25

更新履歴 / Update history
- 2026-01-25: v1.3 ［行列入れ替え］のロジックを実装（歯抜け対応・左上基準）
- 2026-01-25: v1.2 ［行列入れ替え］チェックボックス（UIのみ）を追加、冒頭コメントを更新
- 2025-10-31: v1.1 ローカライズ対応、タイトルにバージョン表示、コメントを日英に整理
- 2025-10-31: v1.0 プレビューの常時ON、連動（上下ディム）、↑↓キーでの増減
*/

//
// バージョン / Version
//
var SCRIPT_VERSION = "v1.3";

//
// 言語判定 / Detect current language
//
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

//
// ラベル定義 / Label definitions
//
var LABELS = {
    dialogTitle: {
        ja: "グリッドの間隔を再定義",
        en: "Redefine Grid Spacing"
    },
    panelSpacing: {
        ja: "間隔",
        en: "Spacing"
    },
    horizontal: {
        ja: "左右:",
        en: "H:"
    },
    vertical: {
        ja: "上下:",
        en: "V:"
    },
    link: {
        ja: "連動",
        en: "Link"
    },
    transpose: {
        ja: "行列入れ替え",
        en: "Swap Rows/Columns"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

/*
ラベルを取得するヘルパー / Helper to get label
*/
function L(key) {
    if (LABELS[key]) {
        return LABELS[key][lang] || LABELS[key].ja || LABELS[key].en || key;
    }
    return key;
}

/* 単位コードとラベルのマップ / Unit code → label map */
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* 現在の単位ラベルを取得 / Get current ruler unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/*
EditTextで↑↓キーによる値の増減を可能にする
Enable arrow-key increment on EditText
- ↑ / ↓ : ±1
- Shift + ↑ / ↓ : ±10 (snap to 10)
- Option(Alt) + ↑ / ↓ : ±0.1
*/
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            // 10単位で増減 / change by 10
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            // 0.1単位で増減 / change by 0.1
            delta = 0.1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            // 1単位 / change by 1
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        // 丸め / rounding
        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = value;

        // 値変更後にプレビュー / update preview after change
        if (typeof updatePreview === "function") {
            updatePreview();
        }
    });
}

/*
ダイアログの生成と表示 / Build & show dialog
*/
function showGridSpacingDialog(applySpacing, runTransposeIfNeeded, restoreOriginalPositions, getCurrentUnitLabel, changeValueByArrowKey, initialGapX) {
    // ここでタイトルとバージョンを合成 / combine title and version here
    var dlgTitle = L('dialogTitle') + ' ' + SCRIPT_VERSION;

    var dlg = new Window('dialog', dlgTitle);
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';

    // パネル名に単位を出す / show unit in panel title
    var pGap = dlg.add('panel', undefined, L('panelSpacing') + ' (' + getCurrentUnitLabel() + ')');
    pGap.orientation = 'row';
    pGap.alignChildren = 'top';
    pGap.margins = [15, 20, 15, 10];

    // 左カラム / left column
    var colLeft = pGap.add('group');
    colLeft.orientation = 'column';
    colLeft.alignChildren = 'left';

    var g1 = colLeft.add('group');
    g1.add('statictext', undefined, L('horizontal'));
    var inputH = g1.add('edittext', undefined, initialGapX.toFixed(1));
    inputH.characters = 6;
    changeValueByArrowKey(inputH);

    var g2 = colLeft.add('group');
    g2.add('statictext', undefined, L('vertical'));
    var inputV = g2.add('edittext', undefined, '20');
    inputV.characters = 6;
    changeValueByArrowKey(inputV);

    // 右カラム（連動）/ right column (link)
    var colRight = pGap.add('group');
    colRight.orientation = 'column';
    colRight.alignChildren = 'center';
    colRight.alignment = ['fill', 'fill'];
    var spacer = colRight.add('statictext', undefined, '');
    spacer.alignment = ['fill', 'fill'];
    var chkLink = colRight.add('checkbox', undefined, L('link'));

    // 行列入れ替え（UIのみ・ロジックは後日）/ Swap rows/columns (UI only; logic later)
    var gTranspose = dlg.add('group');
    gTranspose.orientation = 'row';
    gTranspose.alignChildren = 'left';
    gTranspose.margins = [15, 0, 15, 0];
    var chkTranspose = gTranspose.add('checkbox', undefined, L('transpose'));
    chkTranspose.value = false;

    // 初期状態 / initial state
    chkLink.value = true;
    inputH.active = true;
    inputV.enabled = false;
    inputV.text = inputH.text;

    // ボタン行 / buttons
    var gBtn = dlg.add('group');
    gBtn.alignment = 'right';
    var cancelBtn = gBtn.add('button', undefined, L('cancel'), { name: 'cancel' });
    var okBtn = gBtn.add('button', undefined, L('ok'), { name: 'ok' });

    /*
    プレビュー更新 / update preview
    */
    function updatePreviewImpl() {
        if (chkLink.value) {
            inputV.enabled = false;
            inputV.text = inputH.text;
        } else {
            inputV.enabled = true;
        }

        var gx = parseFloat(inputH.text);
        var gy = parseFloat(inputV.text);
        if (isNaN(gx)) gx = 0;
        if (isNaN(gy)) gy = 0;

        applySpacing(gx, gy);
        if (chkTranspose.value) {
            runTransposeIfNeeded();
        }
    }

    // changeValueByArrowKey() から呼べるようにグローバルへ / expose for arrow-key handler
    updatePreview = updatePreviewImpl;
    $.global.updatePreview = updatePreviewImpl;

    // イベント / events
    inputH.onChanging = function () { updatePreviewImpl(); };
    inputV.onChanging = function () {
        if (!chkLink.value) {
            updatePreviewImpl();
        }
    };
    chkLink.onClick = function () { updatePreviewImpl(); };
    // 行列入れ替え / swap rows & columns
    chkTranspose.onClick = function () { updatePreviewImpl(); };

    // 開いたときに一度プレビュー / first preview when opened
    updatePreviewImpl();

    var result = dlg.show();

    if (result == 1) {
        // OK時は最終値で適用 / apply with final values
        var gx2 = parseFloat(inputH.text);
        var gy2 = parseFloat(inputV.text);
        if (isNaN(gx2)) gx2 = 0;
        if (isNaN(gy2)) gy2 = 0;
        if (chkLink.value) gy2 = gx2;
        applySpacing(gx2, gy2);
        if (chkTranspose.value) {
            runTransposeIfNeeded();
        }
    } else {
        // キャンセル時は元に戻す / restore when canceled
        restoreOriginalPositions();
    }
}

/*
メイン処理 / Main entry
*/
function main() {
    // ドキュメントチェック / document check
    if (app.documents.length === 0) {
        alert((lang === "ja") ? "ドキュメントを開いてください。" : "Open a document first.");
        return;
    }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert((lang === "ja") ? "グリッド状に並んだオブジェクトを選択してください。" : "Please select grid-like objects first.");
        return;
    }

    // 選択を拾う / collect selection
    var items = [];
    for (var i = 0; i < doc.selection.length; i++) {
        var it = doc.selection[i];
        items.push(it);
    }
    if (items.length < 2) {
        alert((lang === "ja") ? "2つ以上のオブジェクトを選択してください。" : "Please select at least two objects.");
        return;
    }

    // Sort selected items by top (Y) descending and left (X) ascending within same row
    var sortedByPosition = items.slice().sort(function(a, b) {
        var ga = a.geometricBounds;
        var gb = b.geometricBounds;
        if (Math.abs(ga[1] - gb[1]) < 1) {
            return ga[0] - gb[0]; // same row → compare left
        }
        return gb[1] - ga[1]; // sort by top descending
    });

    var gb1 = sortedByPosition[0].geometricBounds;
    var gb2 = sortedByPosition[1].geometricBounds;
    var initialGapX = gb2[0] - gb1[2];
    if (initialGapX < 0) initialGapX = 0;

    // 元位置を保存 / save original positions
    var originals = [];
    for (var k = 0; k < items.length; k++) {
        var gb = items[k].geometricBounds;
        originals.push({
            item: items[k],
            left: gb[0],
            top: gb[1]
        });
    }

    /*
    行・列の推定 / build layout info
    */
    function buildLayoutInfo() {
        var bbs = [];
        var minW = Number.MAX_VALUE;
        var minH = Number.MAX_VALUE;

        for (var j = 0; j < items.length; j++) {
            var b = items[j].geometricBounds;
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            if (w < minW) minW = w;
            if (h < minH) minH = h;
            bbs.push({ item: items[j], gb: b });
        }

        // 列まとめ / group columns
        var colCenters = [];
        var colTolerance = minW * 0.5;
        if (colTolerance < 1) colTolerance = 1;

        bbs.sort(function (a, b) { return a.gb[0] - b.gb[0]; });

        for (var c = 0; c < bbs.length; c++) {
            var bb = bbs[c];
            var merged = false;
            var cx = bb.gb[0];
            for (var cc = 0; cc < colCenters.length; cc++) {
                if (Math.abs(colCenters[cc].x - cx) <= colTolerance) {
                    colCenters[cc].items.push(bb);
                    colCenters[cc].x = (colCenters[cc].x * (colCenters[cc].items.length - 1) + cx) / colCenters[cc].items.length;
                    merged = true;
                    break;
                }
            }
            if (!merged) {
                colCenters.push({ x: cx, items: [bb] });
            }
        }

        // 行まとめ / group rows
        var rowCenters = [];
        var rowTolerance = minH * 0.5;
        if (rowTolerance < 1) rowTolerance = 1;

        var bbsY = bbs.slice().sort(function (a, b) { return b.gb[1] - a.gb[1]; });

        for (var r = 0; r < bbsY.length; r++) {
            var bb2 = bbsY[r];
            var merged2 = false;
            var cy = bb2.gb[1];
            for (var rr = 0; rr < rowCenters.length; rr++) {
                if (Math.abs(rowCenters[rr].y - cy) <= rowTolerance) {
                    rowCenters[rr].items.push(bb2);
                    rowCenters[rr].y = (rowCenters[rr].y * (rowCenters[rr].items.length - 1) + cy) / rowCenters[rr].items.length;
                    merged2 = true;
                    break;
                }
            }
            if (!merged2) {
                rowCenters.push({ y: cy, items: [bb2] });
            }
        }

        // 並び順の確定 / sort
        colCenters.sort(function (a, b) { return a.x - b.x; });
        rowCenters.sort(function (a, b) { return b.y - a.y; });

        // 各列の最大幅 / max width per column
        var colWidths = [];
        for (var ci = 0; ci < colCenters.length; ci++) {
            var maxW = 0;
            for (var ci2 = 0; ci2 < colCenters[ci].items.length; ci2++) {
                var gb2 = colCenters[ci].items[ci2].gb;
                var w2 = gb2[2] - gb2[0];
                if (w2 > maxW) maxW = w2;
            }
            colWidths.push(maxW);
        }

        // 各行の最大高さ / max height per row
        var rowHeights = [];
        for (var ri = 0; ri < rowCenters.length; ri++) {
            var maxH = 0;
            for (var ri2 = 0; ri2 < rowCenters[ri].items.length; ri2++) {
                var gb3 = rowCenters[ri].items[ri2].gb;
                var h3 = gb3[1] - gb3[3];
                if (h3 > maxH) maxH = h3;
            }
            rowHeights.push(maxH);
        }

        var baseX = colCenters[0].x;
        var baseY = rowCenters[0].y;

        return {
            bbs: bbs,
            colCenters: colCenters,
            rowCenters: rowCenters,
            colWidths: colWidths,
            rowHeights: rowHeights,
            baseX: baseX,
            baseY: baseY
        };
    }

    var layoutInfo = buildLayoutInfo();

    /*
    行列入れ替え（歯抜け対応）/ Transpose rows & columns (tolerate missing cells)
    - 選択の可視境界（visibleBounds）の left/top を基準に、行・列をクラスタリングして推定
    - 推定したピッチ（隣接差の中央値）で左上基準に再配置
    */
    function transposeGridWithHoles() {
        if (!items || items.length < 1) return;

        // ---- 調整パラメータ / tuning params
        var SNAP_X_TOL = 8.0;   // X方向の「同じ列」とみなす許容（pt）
        var SNAP_Y_TOL = 8.0;   // Y方向の「同じ行」とみなす許容（pt）

        // visibleBounds: [left, top, right, bottom]
        function leftX(it) { return it.visibleBounds[0]; }
        function topY(it) { return it.visibleBounds[1]; }

        // 近い値をクラスタリングして中心値配列を作る / cluster near values into centers
        function clusterValues(values, tol) {
            values.sort(function (a, b) { return a - b; });
            var centers = [];
            for (var i = 0; i < values.length; i++) {
                var v = values[i];
                var found = -1;
                for (var c = 0; c < centers.length; c++) {
                    if (Math.abs(v - centers[c]) <= tol) { found = c; break; }
                }
                if (found < 0) centers.push(v);
                else centers[found] = (centers[found] + v) / 2.0;
            }
            centers.sort(function (a, b) { return a - b; });
            return centers;
        }

        function nearestIndex(arr, v) {
            var best = 0;
            var bestD = Math.abs(v - arr[0]);
            for (var i = 1; i < arr.length; i++) {
                var d = Math.abs(v - arr[i]);
                if (d < bestD) { bestD = d; best = i; }
            }
            return best;
        }

        // 値の差分（隣接）の中央値 / median of adjacent diffs
        function medianDiff(sortedDescOrAsc) {
            if (sortedDescOrAsc.length < 2) return 0;
            var diffs = [];
            for (var i = 1; i < sortedDescOrAsc.length; i++) {
                diffs.push(Math.abs(sortedDescOrAsc[i] - sortedDescOrAsc[i - 1]));
            }
            diffs.sort(function (a, b) { return a - b; });
            return diffs[Math.floor(diffs.length / 2)];
        }

        // left/top を集める / collect left/top
        var xs = [], ys = [];
        for (var k = 0; k < items.length; k++) {
            xs.push(leftX(items[k]));
            ys.push(topY(items[k]));
        }

        var colXs = clusterValues(xs, SNAP_X_TOL); // left -> right
        var rowYs = clusterValues(ys, SNAP_Y_TOL); // will sort top -> bottom next

        // Illustrator座標では上ほどYが大きいことが多いので「上→下」/ sort top -> bottom
        rowYs.sort(function (a, b) { return b - a; });

        // 各オブジェクトを(行,列)に割り当て / assign each item to (row, col)
        var occupancy = {}; // key "r,c" -> item
        var mapping = [];   // {item, r, c}
        for (var m = 0; m < items.length; m++) {
            var it = items[m];
            var r = nearestIndex(rowYs, topY(it));
            var c = nearestIndex(colXs, leftX(it));
            var key = r + "," + c;
            if (occupancy[key]) {
                alert((lang === "ja")
                    ? "同一セルに複数オブジェクトが割り当てられました。\n許容値を下げるか、整列状態を確認してください。\n衝突セル: (" + r + "," + c + ")"
                    : "Multiple objects were assigned to the same cell.\nReduce tolerances or check alignment.\nConflict cell: (" + r + "," + c + ")");
                return;
            }
            occupancy[key] = it;
            mapping.push({ item: it, r: r, c: c });
        }

        var pitchX = medianDiff(colXs);
        var pitchY = medianDiff(rowYs);
        if (pitchX === 0 || pitchY === 0) {
            // 行または列が1つしかない場合は何もしない / nothing meaningful to transpose
            return;
        }

        // 転置後グリッドの基準（左上固定）/ origin at top-left
        var originLeft = colXs[0];
        var originTop = rowYs[0];

        // 転置: newCol = oldRow, newRow = oldCol
        for (var t = 0; t < mapping.length; t++) {
            var obj = mapping[t].item;
            var r0 = mapping[t].r;
            var c0 = mapping[t].c;

            var newCol = r0;
            var newRow = c0;

            var targetLeft = originLeft + newCol * pitchX;
            var targetTop = originTop - newRow * pitchY;

            var b = obj.visibleBounds;
            var curLeft = b[0];
            var curTop = b[1];

            obj.translate(targetLeft - curLeft, targetTop - curTop);
        }

        app.redraw();
    }

    /* 元位置に戻す / restore original positions */
    function restoreOriginalPositions() {
        for (var i2 = 0; i2 < originals.length; i2++) {
            var o = originals[i2];
            var gb4 = o.item.geometricBounds;
            var curLeft = gb4[0];
            var curTop = gb4[1];
            o.item.translate(o.left - curLeft, o.top - curTop);
        }
    }

    /* 間隔を適用 / apply spacing */
    function applySpacing(gapX, gapY) {
        // いったん元に戻す / restore first
        restoreOriginalPositions();

        var bbs = layoutInfo.bbs;
        var colCenters = layoutInfo.colCenters;
        var rowCenters = layoutInfo.rowCenters;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;
        var baseX = layoutInfo.baseX;
        var baseY = layoutInfo.baseY;

        for (var k2 = 0; k2 < bbs.length; k2++) {
            var cur = bbs[k2];

            // 元座標を探す / find original
            var origLeft = null, origTop = null;
            for (var oo = 0; oo < originals.length; oo++) {
                if (originals[oo].item === cur.item) {
                    origLeft = originals[oo].left;
                    origTop = originals[oo].top;
                    break;
                }
            }
            if (origLeft === null) {
                origLeft = cur.gb[0];
                origTop = cur.gb[1];
            }

            // 最も近い列 / nearest column
            var colIndex = 0;
            var minDX = Number.MAX_VALUE;
            for (var c2 = 0; c2 < colCenters.length; c2++) {
                var dx = Math.abs(colCenters[c2].x - cur.gb[0]);
                if (dx < minDX) {
                    minDX = dx;
                    colIndex = c2;
                }
            }

            // 最も近い行 / nearest row
            var rowIndex = 0;
            var minDY = Number.MAX_VALUE;
            for (var r2 = 0; r2 < rowCenters.length; r2++) {
                var dy = Math.abs(rowCenters[r2].y - cur.gb[1]);
                if (dy < minDY) {
                    minDY = dy;
                    rowIndex = r2;
                }
            }

            // 新しいX / new X
            var newX = baseX;
            for (var cc2 = 0; cc2 < colIndex; cc2++) {
                newX += colWidths[cc2] + gapX;
            }

            // 新しいY / new Y
            var newY = baseY;
            for (var rr2 = 0; rr2 < rowIndex; rr2++) {
                newY -= (rowHeights[rr2] + gapY);
            }

            var dxMove = newX - origLeft;
            var dyMove = newY - origTop;
            cur.item.translate(dxMove, dyMove);
        }

        // 再描画 / redraw
        app.redraw();
    }

    // ダイアログ表示 / show dialog
    showGridSpacingDialog(applySpacing, transposeGridWithHoles, restoreOriginalPositions, getCurrentUnitLabel, changeValueByArrowKey, initialGapX);
}

// 実行 / run
main();
