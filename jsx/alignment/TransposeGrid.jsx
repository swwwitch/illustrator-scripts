#target illustrator

/*
### 概要：
選択した複数オブジェクトを行・列として自動判定し、歯抜け（欠け）を許容しつつ、
元の配置ピッチを推定してグリッドを転置（行⇄列）します。
1行だけ→1列、1列だけ→1行の転置にも対応します。
整列されていない選択でも、一定の許容値内で行列を推定し、左上基準で再配置します。

### 更新日：
2026-01-26

v1.1 1行1列に対応
*/

(function () {
    if (app.documents.length === 0) { alert("ドキュメントがありません。"); return; }
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length < 1) { alert("オブジェクトを選択してください。"); return; }

    // ---- 調整パラメータ
    var SNAP_X_TOL = 8.0;   // X方向の「同じ列」とみなす許容（pt）
    var SNAP_Y_TOL = 8.0;   // Y方向の「同じ行」とみなす許容（pt）

    // visibleBounds: [left, top, right, bottom]
    function leftX(it){ return it.visibleBounds[0]; }
    function topY(it){ return it.visibleBounds[1]; }

    // 近い値をクラスタリングして「列X一覧」「行Y一覧」を作る
    function clusterValues(values, tol){
        values.sort(function(a,b){ return a-b; });
        var centers = [];
        for (var i=0; i<values.length; i++){
            var v = values[i];
            var found = -1;
            for (var c=0; c<centers.length; c++){
                if (Math.abs(v - centers[c]) <= tol) { found = c; break; }
            }
            if (found < 0) centers.push(v);
            else {
                // 中心を軽く平均して安定させる
                centers[found] = (centers[found] + v) / 2.0;
            }
        }
        // もう一度ソートして安定化
        centers.sort(function(a,b){ return a-b; });
        return centers;
    }

    function nearestIndex(arr, v){
        var best = 0;
        var bestD = Math.abs(v - arr[0]);
        for (var i=1; i<arr.length; i++){
            var d = Math.abs(v - arr[i]);
            if (d < bestD){ bestD = d; best = i; }
        }
        return best;
    }

    // 選択を配列化
    var items = [];
    for (var i=0; i<sel.length; i++) items.push(sel[i]);

    // すべての left/top を集めて行・列候補を作る
    var xs = [], ys = [];
    for (var k=0; k<items.length; k++){
        xs.push(leftX(items[k]));
        ys.push(topY(items[k]));
    }
    var colXs = clusterValues(xs, SNAP_X_TOL); // 左→右
    var rowYs = clusterValues(ys, SNAP_Y_TOL); // 下→上になりがちなので後で並べ替える

    // Illustrator座標では上ほどYが大きいことが多いので「上→下」に並べる
    rowYs.sort(function(a,b){ return b-a; });

    var rowCount = rowYs.length;
    var colCount = colXs.length;

    // 各オブジェクトを(行,列)に割り当て
    // 歯抜けOK。ただし同一セルに複数来たら警告して止める（必要なら“近い方優先”などに拡張可）
    var occupancy = {}; // key "r,c" -> item
    var mapping = [];   // {item, r, c}
    for (var m=0; m<items.length; m++){
        var it = items[m];
        var r = nearestIndex(rowYs, topY(it));
        var c = nearestIndex(colXs, leftX(it));

        var key = r + "," + c;
        if (occupancy[key]) {
            alert("同一セルに複数オブジェクトが割り当てられました。\n" +
                  "許容値(SNAP_X_TOL / SNAP_Y_TOL)を下げるか、整列状態を確認してください。\n" +
                  "衝突セル: (" + r + "," + c + ")");
            return;
        }
        occupancy[key] = it;
        mapping.push({ item: it, r: r, c: c });
    }

    // 転置後のグリッドは
    //  Xは「行インデックス」を列方向に（ただし rowCount が列数になる）
    //  Yは「列インデックス」を行方向に（ただし colCount が行数になる）
    // つまり転置先の列X一覧＝ rowYsに対応する数だけ必要…ではなく
    // 「元の列X一覧」「元の行Y一覧」をそのまま使って転置すると形が変わるので、
    // 今回は “元の列ピッチ/元の行ピッチ” を使って新しい軸を作ります。

    // 代表ピッチを推定（隣接差の中央値）
    function medianDiff(sortedDescOrAsc){
        if (sortedDescOrAsc.length < 2) return 0;
        var diffs = [];
        for (var i=1; i<sortedDescOrAsc.length; i++){
            diffs.push(Math.abs(sortedDescOrAsc[i] - sortedDescOrAsc[i-1]));
        }
        diffs.sort(function(a,b){ return a-b; });
        return diffs[Math.floor(diffs.length/2)];
    }

    // 元の列/行のピッチを推定（隣接差の中央値）
    // ※行や列が1つしかない場合は0になる
    var pitchX = (colXs.length >= 2) ? medianDiff(colXs) : 0;
    var pitchY = (rowYs.length >= 2) ? medianDiff(rowYs) : 0;

    // 1行→1列、1列→1行にも対応
    // - 1行しかない場合: 横方向ピッチ(pitchX)を縦方向の並び間隔として流用
    // - 1列しかない場合: 縦方向ピッチ(pitchY)を横方向の並び間隔として流用
    var pitchXEff = pitchX;
    var pitchYEff = pitchY;

    if (rowCount === 1 && colCount === 1) {
        alert("1つしか選択されていないため、転置できません。");
        return;
    }

    if (rowCount === 1 && colCount > 1) {
        // 1行 → 1列
        if (pitchX === 0) {
            alert("1行は検出できましたが、列ピッチが推定できませんでした。");
            return;
        }
        pitchXEff = pitchX;   // 新しいX方向は列数=1なので実質使われにくいが、定義しておく
        pitchYEff = pitchX;   // 横ピッチを縦ピッチとして使用
    } else if (colCount === 1 && rowCount > 1) {
        // 1列 → 1行
        if (pitchY === 0) {
            alert("1列は検出できましたが、行ピッチが推定できませんでした。");
            return;
        }
        pitchXEff = pitchY;   // 縦ピッチを横ピッチとして使用
        pitchYEff = pitchY;   // 新しいY方向は行数=1なので実質使われにくいが、定義しておく
    } else {
        // 通常（2行以上 かつ 2列以上）
        if (pitchX === 0 || pitchY === 0) {
            alert("行または列のピッチが推定できませんでした。許容値や整列状態を確認してください。");
            return;
        }
    }

    // 転置後グリッドの基準（左上固定）
    var originLeft = colXs[0];  // 最左
    var originTop  = rowYs[0];  // 最上

    // 転置先の列数=元の行数、行数=元の列数
    // 目標位置（見かけの左上が合うように移動）
    for (var t=0; t<mapping.length; t++){
        var obj = mapping[t].item;
        var r0 = mapping[t].r;
        var c0 = mapping[t].c;

        var newCol = r0;
        var newRow = c0;

        var targetLeft = originLeft + newCol * pitchXEff;
        var targetTop  = originTop  - newRow * pitchYEff;

        var b = obj.visibleBounds;
        var curLeft = b[0];
        var curTop  = b[1];

        obj.translate(targetLeft - curLeft, targetTop - curTop);
    }

    // alert("転置しました（歯抜け対応）。\n" +
    //       "検出: 行=" + rowCount + " 列=" + colCount + "\n" +
    //       "推定ピッチ: X=" + pitchX.toFixed(2) + "pt, Y=" + pitchY.toFixed(2) + "pt");
})();
