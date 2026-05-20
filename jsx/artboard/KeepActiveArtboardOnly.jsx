#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

// 矩形の交差判定（[top,left,bottom,right]）/ Rectangle intersection test
function rectanglesIntersect(rectA, rectB) {
    // 右 < 左 / left > right / 上 < 下 のいずれかで非交差
    if (rectA[3] < rectB[1]) return false; // a.right < b.left
    if (rectA[1] > rectB[3]) return false; // a.left > b.right
    if (rectA[0] < rectB[2]) return false; // a.top < b.bottom
    if (rectA[2] > rectB[0]) return false; // a.bottom > b.top
    return true;
}

// artboardRect([left, top, right, bottom]) → [top, left, bottom, right]
function artboardRectToBounds(artboardRect) {
    return [artboardRect[1], artboardRect[0], artboardRect[3], artboardRect[2]];
}

// 点が矩形内にあるか（含む境界）/ Point-in-rect (inclusive)
function isPointInRect(x, y, rect) { // rect: [left, top, right, bottom]
    return (x >= rect[0] && x <= rect[2] && y <= rect[1] && y >= rect[3]);
}

// 2Dベクトル外積 / Cross product helper
function cross2D(ax, ay, bx, by) {
    return ax * by - ay * bx;
}

// 線分同士の交差判定（座標）/ Segment-segment intersection
function segmentsIntersect(p1x, p1y, p2x, p2y, q1x, q1y, q2x, q2y) {
    var pDx = p2x - p1x,
        pDy = p2y - p1y;
    var qDx = q2x - q1x,
        qDy = q2y - q1y;
    var denominator = cross2D(pDx, pDy, qDx, qDy);
    var p1ToQ1Dx = q1x - p1x,
        p1ToQ1Dy = q1y - p1y;
    if (denominator === 0) {
        // 平行（含む共線）/ Parallel (including collinear)
        // 共線の場合、投影で重なりチェック（簡易）/ If collinear, check overlap via projection
        var collinearCross = cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy);
        if (Math.abs(collinearCross) > 1e-6) return false; // 非共線 / not collinear
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
    var tParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, qDx, qDy) / denominator;
    var uParam = cross2D(p1ToQ1Dx, p1ToQ1Dy, pDx, pDy) / denominator;
    return (tParam >= 0 && tParam <= 1 && uParam >= 0 && uParam <= 1);
}

// 線分と矩形の交差判定 / Segment-rectangle intersection
function segmentIntersectsRect(x1, y1, x2, y2, rect) { // rect: [left, top, right, bottom]
    // 端点のどちらかが矩形内
    if (isPointInRect(x1, y1, rect) || isPointInRect(x2, y2, rect)) return true;
    // 矩形の4辺
    var left = rect[0],
        top = rect[1],
        right = rect[2],
        bottom = rect[3];
    // 上辺 (left,top)-(right,top)
    if (segmentsIntersect(x1, y1, x2, y2, left, top, right, top)) return true;
    // 下辺 (left,bottom)-(right,bottom)
    if (segmentsIntersect(x1, y1, x2, y2, left, bottom, right, bottom)) return true;
    // 左辺 (left,top)-(left,bottom)
    if (segmentsIntersect(x1, y1, x2, y2, left, top, left, bottom)) return true;
    // 右辺 (right,top)-(right,bottom)
    if (segmentsIntersect(x1, y1, x2, y2, right, top, right, bottom)) return true;
    return false;
}

// パスがアートボード矩形に重なるか（端点内包 or 線分交差）/ Does path overlap artboard rect?
function pathIntersectsArtboard(pathItem, artboardRect) { // artboardRect: [left, top, right, bottom]
    try {
        var pathPoints = pathItem.pathPoints;
        if (!pathPoints || pathPoints.length === 0) return false;
        for (var i = 0; i < pathPoints.length; i++) {
            var currentAnchor = pathPoints[i].anchor; // [x, y]
            var nextAnchor = pathPoints[(i + 1) % pathPoints.length].anchor; // next (wrap if closed)
            // 開パスで最後のセグメントは結ばない / For open paths, skip last-to-first
            if (!pathItem.closed && i === pathPoints.length - 1) break;
            if (segmentIntersectsRect(currentAnchor[0], currentAnchor[1], nextAnchor[0], nextAnchor[1], artboardRect)) return true;
        }
        return false;
    } catch (e) {
        return false;
    }
}

// レイヤー/グループの状態を再帰的に保存し一時的に解除するユーティリティ
// Utilities to recursively capture and relax layer/group states
function collectLayerStates(layer, stateLog) {
    // 元状態を記録 / Record original state
    stateLog.push({
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
        collectLayerStates(layer.layers[i], stateLog);
    }
    // グループ / Groups within this layer
    for (var j = 0; j < layer.groupItems.length; j++) {
        collectGroupStates(layer.groupItems[j], stateLog);
    }
    // レイヤー直下アイテム / Items directly under this layer
    collectItemStatesInLayer(layer, stateLog);
}

function collectGroupStates(groupItem, stateLog) {
    stateLog.push({
        type: 'Group',
        ref: groupItem,
        locked: groupItem.locked,
        visible: groupItem.visible
    });
    groupItem.locked = false;
    groupItem.visible = true;
    // ネストグループ / Nested groups
    for (var i = 0; i < groupItem.groupItems.length; i++) {
        collectGroupStates(groupItem.groupItems[i], stateLog);
    }
    // グループ直下アイテム / Items directly under this group
    collectItemStatesInGroup(groupItem, stateLog);
}

// レイヤー配下の PageItem の状態を保存し一時解除 / Collect & relax PageItems under a layer
function collectItemStatesInLayer(layer, stateLog) {
    try {
        var items = layer.pageItems;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            stateLog.push({
                type: 'Item',
                ref: item,
                locked: item.locked,
                hidden: item.hidden
            });
            // 一時解除 / Relax temporarily
            if (item.locked) item.locked = false;
            if (item.hidden) item.hidden = false;
        }
    } catch (e) {}
}

// グループ配下の PageItem の状態を保存し一時解除 / Collect & relax PageItems under a group
function collectItemStatesInGroup(groupItem, stateLog) {
    try {
        var items = groupItem.pageItems;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            stateLog.push({
                type: 'Item',
                ref: item,
                locked: item.locked,
                hidden: item.hidden
            });
            if (item.locked) item.locked = false;
            if (item.hidden) item.hidden = false;
        }
    } catch (e) {}
}

function captureAndUnlockStructure(doc) {
    var stateLog = [];
    for (var i = 0; i < doc.layers.length; i++) {
        collectLayerStates(doc.layers[i], stateLog);
    }
    return stateLog;
}

function restoreStructure(savedStates) {
    // 逆順で復元（下位→上位）/ Restore in reverse (bottom-up)
    for (var i = savedStates.length - 1; i >= 0; i--) {
        var state = savedStates[i];
        try {
            if (state.type === 'Layer') {
                state.ref.locked = state.locked;
                state.ref.visible = state.visible;
                state.ref.template = state.template;
            } else if (state.type === 'Group') {
                state.ref.locked = state.locked;
                state.ref.visible = state.visible;
            } else if (state.type === 'Item') {
                state.ref.locked = state.locked;
                state.ref.hidden = state.hidden;
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
    var savedStructureStates = captureAndUnlockStructure(doc);

    // アクティブArtboardと重ならないガイド（オブジェクトガイド）を削除（バウンディングではなくパスの実体で判定）
    // Delete object guides that do NOT overlap the active artboard (check path segments, not only bounds)
    try {
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var artboardRect = doc.artboards[activeArtboardIndex].artboardRect; // [left, top, right, bottom]

        var guidesToRemove = [];
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.guides === true) {
                var intersectsArtboard = false;
                // CompoundPathItem / PathItem の両方に対応
                if (item.typename === 'CompoundPathItem') {
                    var subPaths = item.pathItems;
                    for (var j = 0; j < subPaths.length && !intersectsArtboard; j++) {
                        if (pathIntersectsArtboard(subPaths[j], artboardRect)) intersectsArtboard = true;
                    }
                } else if (item.typename === 'PathItem') {
                    if (pathIntersectsArtboard(item, artboardRect)) intersectsArtboard = true;
                } else {
                    // その他はバウンディングでフォールバック
                    var geometricBounds = item.geometricBounds; // [top, left, bottom, right]
                    var artboardBounds = [artboardRect[0], artboardRect[1], artboardRect[2], artboardRect[3]]; // [left, top, right, bottom]
                    // geometricBounds [top,left,bottom,right] を [left,top,right,bottom] へ変換
                    var boundsRect = [geometricBounds[1], geometricBounds[0], geometricBounds[3], geometricBounds[2]];
                    // 簡易交差: 端点内包なしの矩形交差
                    intersectsArtboard = !(boundsRect[2] < artboardBounds[0] || boundsRect[0] > artboardBounds[2] || boundsRect[1] < artboardBounds[3] || boundsRect[3] > artboardBounds[1]);
                }
                if (!intersectsArtboard) guidesToRemove.push(item);
            }
        }
        for (var k = 0; k < guidesToRemove.length; k++) guidesToRemove[k].remove();
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
    restoreStructure(savedStructureStates);
}

main();

})();
