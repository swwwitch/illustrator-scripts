#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

RemoveOtherArtboards

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RemoveOtherArtboards.jsx

### 概要：

- アクティブなアートボード以外を削除
- アクティブなアートボード外のオブジェクト削除（ガイドを含む）

### 主な機能：

- アクティブ以外のアートボードを削除
- アクティブArtboard外のオブジェクト削除
- レイヤー/サブレイヤー/グループのロック・表示・テンプレート状態の一時解除→復元
- アクティブArtboardと重ならないオブジェクトガイドの削除（ルーラーガイドは対象外）
- PageItemのlocked/hidden一時解除→復元

### 処理の流れ：

1) アクティブドキュメント・アートボード取得
2) アクティブ以外のアートボードを末尾から削除
3) レイヤー/サブレイヤー/グループ/PageItemの状態を再帰的に収集し一時解除
4) アクティブArtboardと重ならないオブジェクトガイドを削除
5) 反転選択でアクティブArtboard外オブジェクトを削除
6) 状態を復元

### 更新履歴：

- v1.0 (202406XX) : 初期バージョン
- v1.1 (20250815) : 説明文強化、ロック復元・ガイド削除ロジックの注記、コメント整理
- v1.2 (20250815) : PageItemのlocked/hidden対応、用語統一、構造化コメント

---

### Script Name:

RemoveOtherArtboards

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RemoveOtherArtboards.jsx

### Description:

- Remove all non-active artboards
- Delete objects outside the active artboard (including guides)

### Key Features:

- Delete non-active artboards
- Delete objects outside the active artboard
- Temporarily relax and restore lock/visible/template state of layers/sublayers/groups
- Delete object guides not overlapping the active artboard (ruler guides excluded)
- Temporarily relax and restore locked/hidden state of PageItems

### Processing Flow:

1) Get the active document and active artboard
2) Remove non-active artboards from the end
3) Recursively collect and temporarily relax the state of layers/sublayers/groups/PageItems
4) Delete object guides not overlapping the active artboard
5) Inverse-select and delete objects outside the active artboard
6) Restore original states

### Update History:

- v1.0 (202406XX) : Initial version
- v1.1 (20250815) : Expanded docs, notes on lock restore & guide checks, comment cleanup
- v1.2 (20250815) : Added PageItem locked/hidden handling, terminology unification, structured comments
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

// 矩形の交差判定（[top,left,bottom,right]）/ Rectangle intersection test
function rectIntersects(a, b) {
    // 右 < 左 / left > right / 上 < 下 のいずれかで非交差
    if (a[3] < b[1]) return false; // a.right < b.left
    if (a[1] > b[3]) return false; // a.left > b.right
    if (a[0] < b[2]) return false; // a.top < b.bottom
    if (a[2] > b[0]) return false; // a.bottom > b.top
    return true;
}

// artboardRect([left, top, right, bottom]) → [top, left, bottom, right]
function artboardRectToBounds(ar) {
    return [ar[1], ar[0], ar[3], ar[2]];
}

// 点が矩形内にあるか（含む境界）/ Point-in-rect (inclusive)
function pointInRect(x, y, rect) { // rect: [left, top, right, bottom]
    return (x >= rect[0] && x <= rect[2] && y <= rect[1] && y >= rect[3]);
}

// 2Dベクトル外積 / Cross product helper
function cross(ax, ay, bx, by) {
    return ax * by - ay * bx;
}

// 線分同士の交差判定（座標）/ Segment-segment intersection
function segmentsIntersect(p1x, p1y, p2x, p2y, q1x, q1y, q2x, q2y) {
    var r1x = p2x - p1x,
        r1y = p2y - p1y;
    var r2x = q2x - q1x,
        r2y = q2y - q1y;
    var d = cross(r1x, r1y, r2x, r2y);
    var qp1x = q1x - p1x,
        qp1y = q1y - p1y;
    if (d === 0) {
        // 平行（含む共線）/ Parallel (including collinear)
        // 共線の場合、投影で重なりチェック（簡易）/ If collinear, check overlap via projection
        var c = cross(qp1x, qp1y, r1x, r1y);
        if (Math.abs(c) > 1e-6) return false; // 非共線 / not collinear
        // 軸に沿って重なり判定 / overlap on x/y
        var minPx = Math.min(p1x, p2x),
            maxPx = Math.max(p1x, p2x);
        var minQx = Math.min(q1x, q2x),
            maxQx = Math.max(q1x, q2x);
        var minPy = Math.min(p1y, p2y),
            maxPy = Math.max(p1y, p2y);
        var minQy = Math.min(q1y, q2y),
            maxQy = Math.max(q1y, q2y);
        var overlapX = !(maxPx < minQx || maxQx < minPx);
        var overlapY = !(maxPy < minQy || maxQy < minPy);
        return overlapX && overlapY;
    }
    var t = cross(qp1x, qp1y, r2x, r2y) / d;
    var u = cross(qp1x, qp1y, r1x, r1y) / d;
    return (t >= 0 && t <= 1 && u >= 0 && u <= 1);
}

// 線分と矩形の交差判定 / Segment-rectangle intersection
function segmentIntersectsRect(x1, y1, x2, y2, rect) { // rect: [left, top, right, bottom]
    // 端点のどちらかが矩形内
    if (pointInRect(x1, y1, rect) || pointInRect(x2, y2, rect)) return true;
    // 矩形の4辺
    var L = rect[0],
        T = rect[1],
        R = rect[2],
        B = rect[3];
    // 上辺 (L,T)-(R,T)
    if (segmentsIntersect(x1, y1, x2, y2, L, T, R, T)) return true;
    // 下辺 (L,B)-(R,B)
    if (segmentsIntersect(x1, y1, x2, y2, L, B, R, B)) return true;
    // 左辺 (L,T)-(L,B)
    if (segmentsIntersect(x1, y1, x2, y2, L, T, L, B)) return true;
    // 右辺 (R,T)-(R,B)
    if (segmentsIntersect(x1, y1, x2, y2, R, T, R, B)) return true;
    return false;
}

// パスがアートボード矩形に重なるか（端点内包 or 線分交差）/ Does path overlap artboard rect?
function pathIntersectsArtboard(pathItem, abRect) { // abRect: [left, top, right, bottom]
    try {
        var pts = pathItem.pathPoints;
        if (!pts || pts.length === 0) return false;
        for (var i = 0; i < pts.length; i++) {
            var p = pts[i].anchor; // [x, y]
            var q = pts[(i + 1) % pts.length].anchor; // next (wrap if closed)
            // 開パスで最後のセグメントは結ばない / For open paths, skip last-to-first
            if (!pathItem.closed && i === pts.length - 1) break;
            if (segmentIntersectsRect(p[0], p[1], q[0], q[1], abRect)) return true;
        }
        return false;
    } catch (e) {
        return false;
    }
}

// レイヤー/グループの状態を再帰的に保存し一時的に解除するユーティリティ
// Utilities to recursively capture and relax layer/group states
function collectLayerStates(layer, out) {
    // 元状態を記録 / Record original state
    out.push({
        type: 'Layer',
        ref: layer,
        locked: layer.locked,
        visible: layer.visible,
        template: layer.template
    });
    // 一時解除 / Relax temporarily
    layer.locked = false;
    layer.visible = true;
    if (layer.template === true) layer.template = false;
    // サブレイヤー / Sublayers
    for (var i = 0; i < layer.layers.length; i++) {
        collectLayerStates(layer.layers[i], out);
    }
    // グループ / Groups within this layer
    for (var g = 0; g < layer.groupItems.length; g++) {
        collectGroupStates(layer.groupItems[g], out);
    }
    // レイヤー直下アイテム / Items directly under this layer
    collectItemStatesInLayer(layer, out);
}

function collectGroupStates(groupItem, out) {
    out.push({
        type: 'Group',
        ref: groupItem,
        locked: groupItem.locked,
        visible: groupItem.visible
    });
    groupItem.locked = false;
    groupItem.visible = true;
    // ネストグループ / Nested groups
    for (var i = 0; i < groupItem.groupItems.length; i++) {
        collectGroupStates(groupItem.groupItems[i], out);
    }
    // グループ直下アイテム / Items directly under this group
    collectItemStatesInGroup(groupItem, out);
}

// レイヤー配下の PageItem の状態を保存し一時解除 / Collect & relax PageItems under a layer
function collectItemStatesInLayer(layer, out) {
    try {
        var items = layer.pageItems;
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            out.push({
                type: 'Item',
                ref: it,
                locked: it.locked,
                hidden: it.hidden
            });
            // 一時解除 / Relax temporarily
            if (it.locked) it.locked = false;
            if (it.hidden) it.hidden = false;
        }
    } catch (e) {}
}

// グループ配下の PageItem の状態を保存し一時解除 / Collect & relax PageItems under a group
function collectItemStatesInGroup(groupItem, out) {
    try {
        var items = groupItem.pageItems;
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            out.push({
                type: 'Item',
                ref: it,
                locked: it.locked,
                hidden: it.hidden
            });
            if (it.locked) it.locked = false;
            if (it.hidden) it.hidden = false;
        }
    } catch (e) {}
}

function relaxStructure(doc) {
    var states = [];
    for (var i = 0; i < doc.layers.length; i++) {
        collectLayerStates(doc.layers[i], states);
    }
    return states;
}

function restoreStructure(states) {
    // 逆順で復元（下位→上位）/ Restore in reverse (bottom-up)
    for (var i = states.length - 1; i >= 0; i--) {
        var s = states[i];
        try {
            if (s.type === 'Layer') {
                s.ref.locked = s.locked;
                s.ref.visible = s.visible;
                s.ref.template = s.template;
            } else if (s.type === 'Group') {
                s.ref.locked = s.locked;
                s.ref.visible = s.visible;
            } else if (s.type === 'Item') {
                s.ref.locked = s.locked;
                s.ref.hidden = s.hidden;
            }
        } catch (e) {}
    }
}

function main() {
    // アクティブなドキュメントを取得 / Get the active document
    var doc = app.activeDocument;

    // アートボードの総数を取得 / Get total number of artboards
    var totalArtboards = doc.artboards.length;

    // アートボードが1つだけの場合は何もしない / Do nothing if only one artboard exists
    if (totalArtboards > 1) {
        // アクティブなアートボードのインデックスを取得 / Get index of the active artboard
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

        // 最後からアートボードを削除（アクティブなアートボード以外） / Remove artboards from the end, except the active one
        for (var i = totalArtboards - 1; i >= 0; i--) {
            if (i !== activeArtboardIndex) {
                doc.artboards.remove(i);
            }
        }
    }
    // レイヤー/サブレイヤー/グループのロック・表示・テンプレートを再帰的に一時解除 / Recursively relax lock/visible/template on layers, sublayers, and groups
    var structureStates = relaxStructure(doc);

    // アクティブArtboardと重ならないガイド（オブジェクトガイド）を削除（バウンディングではなくパスの実体で判定）
    // Delete object guides that do NOT overlap the active artboard (check path segments, not only bounds)
    try {
        var abIdx = doc.artboards.getActiveArtboardIndex();
        var abRect = doc.artboards[abIdx].artboardRect; // [left, top, right, bottom]

        var toRemoveGuides = [];
        for (var p = doc.pageItems.length - 1; p >= 0; p--) {
            var it = doc.pageItems[p];
            if (it.guides === true) {
                var keep = false;
                // CompoundPathItem / PathItem の両方に対応
                if (it.typename === 'CompoundPathItem') {
                    var subs = it.pathItems;
                    for (var si = 0; si < subs.length && !keep; si++) {
                        if (pathIntersectsArtboard(subs[si], abRect)) keep = true;
                    }
                } else if (it.typename === 'PathItem') {
                    if (pathIntersectsArtboard(it, abRect)) keep = true;
                } else {
                    // その他はバウンディングでフォールバック
                    var gb = it.geometricBounds; // [top, left, bottom, right]
                    var abBounds = [abRect[0], abRect[1], abRect[2], abRect[3]]; // [left, top, right, bottom]
                    // gb -> [top,left,bottom,right] を [left,top,right,bottom] へ変換
                    var gbRect = [gb[1], gb[0], gb[3], gb[2]];
                    // 簡易交差: 端点内包なしの矩形交差
                    keep = !(gbRect[2] < abBounds[0] || gbRect[0] > abBounds[2] || gbRect[1] < abBounds[3] || gbRect[3] > abBounds[1]);
                }
                if (!keep) toRemoveGuides.push(it);
            }
        }
        for (var r = 0; r < toRemoveGuides.length; r++) toRemoveGuides[r].remove();
    } catch (e) {
        // 可能な範囲で実施（ルーラーガイドは対象外）/ Best-effort (ruler guides excluded)
    }

    // アクティブなアートボード上の全オブジェクトを選択→反転→削除
    // Select all on active artboard, invert selection, then delete
    app.executeMenuCommand('selectallinartboard');
    app.executeMenuCommand('Inverse menu item');
    app.executeMenuCommand('clear');

    // 元々ロックされていたレイヤーを再ロック / Re-lock layers that were originally locked
    // レイヤー/サブレイヤー/グループの状態を復元 / Restore layers/sublayers/groups state
    restoreStructure(structureStates);
}

main();