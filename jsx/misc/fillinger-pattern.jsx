#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
Script: TileSmallIntoLarge.jsx
Purpose (JP):
  2つのオブジェクトを選択し、面積が大きい方を「容器」、小さい方を「タイル」とみなし、
  容器のバウンディングボックスをタイルで均等に敷き詰めます（グリッド配置）。
  結果はマスクなしで、容器の内側に収まるタイルのみ残します。元の2オブジェクトは残します。
Purpose (EN):
  With two selected objects, treats the larger one as the container and the smaller one as the tile.
  Fills the container's bounding box with an even grid of tile duplicates, then keeps only tiles fully inside the container (no mask).
  The two original items remain untouched.

Usage (JP):
  ・ちょうど2つのオブジェクトを選択して実行してください。
  ・ダイアログの単位は Illustrator のルーラー設定（rulerType）に追従します。
  結果はマスクなしで、容器の内側に収まるタイルのみ残します。元の2オブジェクトは残します。
Usage (EN):
  ・Select exactly two objects and run the script.
  ・Spacing units follow Illustrator's current ruler units (rulerType).
  The result keeps only tiles fully inside the container (no mask). The two original items remain.

Update History / 更新履歴:
  - 2025-10-26 v1.0.0 初版 / Initial release.
*/

var SCRIPT_VERSION = "v1.1.0"; // バージョン

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

function L(key) {
    var lang = getCurrentLang();
    try {
        return LABELS[key][lang];
    } catch (e) {
        return key;
    }
}

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        // タイトル末尾にバージョンを表示
        ja: "敷き詰め設定 " + SCRIPT_VERSION,
        en: "Tile Fill Settings " + SCRIPT_VERSION
    }
};

/* タイルグリッド共通ユーティリティ / Tile Grid Common Utilities */
// ---- Namespace: TG (Tile Grid) common utilities ----
var TG = {
    boundsInfo: function(item) {
        var b = item.visibleBounds; // [l,t,r,b]
        var w = b[2] - b[0];
        var h = b[1] - b[3];
        return {
            left: b[0],
            top: b[1],
            right: b[2],
            bottom: b[3],
            width: w,
            height: h,
            area: Math.abs(w * h)
        };
    },
    unitLabelMap: {
        0: 'in',
        1: 'mm',
        2: 'pt',
        3: 'pica',
        4: 'cm',
        5: 'Q/H',
        6: 'px',
        7: 'ft/in',
        8: 'm',
        9: 'yd',
        10: 'ft'
    },
    getRulerType: function() {
        return app.preferences.getIntegerPreference('rulerType');
    },
    getCurrentUnitLabel: function() {
        var unitCode = TG.getRulerType();
        return TG.unitLabelMap[unitCode] || 'pt';
    },
    toPoints: function(val, unitCode) {
        if (val === 0) return 0;
        switch (unitCode) {
            case 0:
                return val * 72.0; // in
            case 1:
                return val * (72.0 / 25.4); // mm
            case 2:
                return val; // pt
            case 3:
                return val * 12.0; // pica
            case 4:
                return val * (72.0 / 2.54); // cm
            case 5:
                return val * (72.0 / 25.4) * 0.25; // Q/H (1Q=0.25mm)
            case 6:
                return val * 1.0; // px ≈ pt (Illustrator geometry)
            case 7:
                return val * 72.0; // ft/in -> treat as inch
            case 8:
                return val * (72.0 / 0.0254); // m
            case 9:
                return val * (72.0 * 36.0); // yd
            case 10:
                return val * (72.0 * 12.0); // ft
            default:
                return val;
        }
    },
    isRectanglePath: function(p) {
        if (p.typename !== 'PathItem' || !p.closed || p.pathPoints.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            if (p.pathPoints[i].pointType !== PointType.CORNER) return false;
        }
        return true;
    },
    isEllipseLike: function(p) {
        if (p.typename !== 'PathItem' || !p.closed || p.pathPoints.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            if (p.pathPoints[i].pointType !== PointType.SMOOTH) return false;
        }
        return true;
    },
    // --- Flatten PathItem into polyline points (document coords) / パスをポリラインに近似 ---
    flattenPath: function(pathItem, stepsPerCurve) {
        var pts = [];
        if (!pathItem || pathItem.typename !== 'PathItem' || !pathItem.closed) return pts;
        var n = pathItem.pathPoints.length;
        if (n < 2) return pts;
        var steps = Math.max(2, stepsPerCurve || 12);
        function lerp(a,b,t){ return [a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t]; }
        function cubic(p0,p1,p2,p3,t){
            var a = lerp(p0,p1,t), b = lerp(p1,p2,t), c = lerp(p2,p3,t);
            var d = lerp(a,b,t), e = lerp(b,c,t); return lerp(d,e,t);
        }
        for (var i=0;i<n;i++){
            var pt = pathItem.pathPoints[i];
            var next = pathItem.pathPoints[(i+1)%n];
            var p0 = [pt.anchor[0], pt.anchor[1]];
            var p1 = [pt.rightDirection[0], pt.rightDirection[1]];
            var p2 = [next.leftDirection[0], next.leftDirection[1]];
            var p3 = [next.anchor[0], next.anchor[1]];
            for (var s=0; s<=steps; s++){
                var t = s/steps; var xy = cubic(p0,p1,p2,p3,t);
                if (i===0 && s===0) pts.push(xy); else if (s>0) pts.push(xy);
            }
        }
        return pts;
    },
    // --- Point in polygon (ray casting) / ポリゴン内点判定 ---
    pointInPolygon: function(x,y, poly){
        var inside=false, j=poly.length-1;
        for (var i=0;i<poly.length;i++){
            var xi=poly[i][0], yi=poly[i][1];
            var xj=poly[j][0], yj=poly[j][1];
            var inter = ((yi>y)!=(yj>y)) && (x < (xj-xi)*(y-yi)/( (yj-yi)||1e-9 ) + xi);
            if (inter) inside = !inside;
            j=i;
        }
        return inside;
    },
    // --- Minimum distance from point to polygon edges / ポリライン最近距離 ---
    minDistToPolyline: function(x,y, poly){
        function dist2(ax,ay,bx,by){ var dx=ax-bx, dy=ay-by; return dx*dx+dy*dy; }
        function segDist(ax,ay,bx,by,px,py){
            var vx=bx-ax, vy=by-ay, wx=px-ax, wy=py-ay;
            var c1= vx*wx+vy*wy; if (c1<=0) return Math.sqrt(dist2(px,py,ax,ay));
            var c2= vx*vx+vy*vy; if (c2<=0) return Math.sqrt(dist2(px,py,ax,ay));
            var t=c1/c2; if (t>=1) return Math.sqrt(dist2(px,py,bx,by));
            var qx=ax+vx*t, qy=ay+vy*t; return Math.sqrt(dist2(px,py,qx,qy));
        }
        var md=1e9, j=poly.length-1;
        for (var i=0;i<poly.length;i++){
            var a=poly[j], b=poly[i];
            var d=segDist(a[0],a[1],b[0],b[1], x,y); if (d<md) md=d; j=i;
        }
        return md;
    },
    // --- Segment intersection (strict) / 線分交差判定 ---
    segIntersect: function(ax,ay,bx,by, cx,cy,dx,dy){
        function orient(px,py,qx,qy, rx,ry){ return (qx-px)*(ry-py) - (qy-py)*(rx-px); }
        function onSeg(px,py,qx,qy, rx,ry){
            return Math.min(px,qx) - 1e-9 <= rx && rx <= Math.max(px,qx) + 1e-9 &&
                   Math.min(py,qy) - 1e-9 <= ry && ry <= Math.max(py,qy) + 1e-9 &&
                   Math.abs(orient(px,py,qx,qy,rx,ry)) < 1e-9;
        }
        var o1 = orient(ax,ay,bx,by, cx,cy);
        var o2 = orient(ax,ay,bx,by, dx,dy);
        var o3 = orient(cx,cy,dx,dy, ax,ay);
        var o4 = orient(cx,cy,dx,dy, bx,by);
        if (((o1>0 && o2<0)||(o1<0 && o2>0)) && ((o3>0 && o4<0)||(o3<0 && o4>0))) return true;
        if (onSeg(ax,ay,bx,by, cx,cy)) return true;
        if (onSeg(ax,ay,bx,by, dx,dy)) return true;
        if (onSeg(cx,cy,dx,dy, ax,ay)) return true;
        if (onSeg(cx,cy,dx,dy, bx,by)) return true;
        return false;
    },
    // --- Rect fully inside polygon: corners inside + no edge intersections / 矩形が多角形に完全内包 ---
    rectInsidePolygon: function(ib, poly, margin){
        if (!poly || poly.length < 3) return false;
        var m = (margin>0? margin: 0);
        var l = ib.left + m;
        var r = ib.right - m;
        var t = ib.top - m;
        var b = ib.bottom + m;
        if (r <= l || t <= b) return false;
        var corners = [ [l,t],[r,t],[l,b],[r,b] ];
        for (var i=0;i<4;i++){
            if (!TG.pointInPolygon(corners[i][0], corners[i][1], poly)) return false;
        }
        var rectEdges = [ [l,t,r,t],[r,t,r,b],[r,b,l,b],[l,b,l,t] ];
        var j = poly.length-1;
        for (var i=0;i<poly.length;i++){
            var a = poly[j], bpt = poly[i];
            for (var e=0;e<4;e++){
                var E = rectEdges[e];
                if (TG.segIntersect(a[0],a[1], bpt[0],bpt[1], E[0],E[1],E[2],E[3])) return false;
            }
            j=i;
        }
        return true;
    }
    ,
    // --- Flatten container (PathItem or CompoundPathItem) into array of polygons / 複合パス対応ポリゴン配列 ---
    flattenContainerPolys: function(container, stepsPerCurve) {
        var polys = [];
        if (!container) return polys;
        if (container.typename === 'CompoundPathItem') {
            for (var i = 0; i < container.pathItems.length; i++) {
                var pi = container.pathItems[i];
                if (!pi.closed) continue;
                var poly = TG.flattenPath(pi, stepsPerCurve || 12);
                if (poly.length >= 3) polys.push(poly);
            }
        } else if (container.typename === 'PathItem') {
            if (container.closed) {
                var poly1 = TG.flattenPath(container, stepsPerCurve || 12);
                if (poly1.length >= 3) polys.push(poly1);
            }
        }
        return polys; // array of [ [x,y], ... ]
    },
    // --- Even-odd fill point test across multiple polygons / 複数ポリゴンの偶奇判定 ---
    pointInPolys: function(x, y, polys) {
        var count = 0;
        for (var i = 0; i < polys.length; i++) {
            if (TG.pointInPolygon(x, y, polys[i])) count++;
        }
        return (count % 2) === 1; // even-odd rule
    },
    // --- Rect fully inside multiple polygons using even-odd rule + no edge intersections ---
    rectInsidePolys: function(ib, polys, margin) {
        if (!polys || polys.length === 0) return false;
        var m = (margin>0? margin: 0);
        var l = ib.left + m, r = ib.right - m, t = ib.top - m, b = ib.bottom + m;
        if (r <= l || t <= b) return false;
        // corners inside even-odd fill
        var corners = [ [l,t],[r,t],[l,b],[r,b] ];
        for (var i=0;i<4;i++) {
            if (!TG.pointInPolys(corners[i][0], corners[i][1], polys)) return false;
        }
        // no rect edge intersects any polygon edge (treat each poly separately)
        var rectEdges = [ [l,t,r,t],[r,t,r,b],[r,b,l,b],[l,b,l,t] ];
        for (var p=0; p<polys.length; p++) {
            var poly = polys[p];
            var j = poly.length-1;
            for (var k=0; k<poly.length; k++) {
                var a = poly[j], bpt = poly[k];
                for (var e=0; e<4; e++) {
                    var E = rectEdges[e];
                    if (TG.segIntersect(a[0],a[1], bpt[0],bpt[1], E[0],E[1],E[2],E[3])) return false;
                }
                j = k;
            }
        }
        return true;
    }
};

function main() {
    var doc = app.documents.length ? app.activeDocument : null;
    if (!doc) {
        alert('ドキュメントが開いていません / No active document.');
        return;
    }

    if (!doc.selection || doc.selection.length !== 2) {
        alert('ちょうど2つを選択してください。/ Please select exactly two objects.');
        return;
    }

    // Identify larger (container) and smaller (tile)
    var a = doc.selection[0];
    var b = doc.selection[1];
    var infoA = TG.boundsInfo(a);
    var infoB = TG.boundsInfo(b);
    var container = infoA.area >= infoB.area ? a : b;
    var tile = (container === a) ? b : a;
    var cInfo = TG.boundsInfo(container);
    var tInfo = TG.boundsInfo(tile);

    if (tInfo.width <= 0 || tInfo.height <= 0) {
        alert('タイルのサイズが不正です / Invalid tile size.');
        return;
    }

    // --- Preview helpers ---
    var _previewGroup = null;

    function clearPreview() {
        // Remove the tracked preview group
        try {
            if (_previewGroup && _previewGroup.isValid) {
                _previewGroup.remove();
            }
        } catch (e1) {}
        _previewGroup = null;

        // Also sweep any stray preview groups by name on the same layer
        try {
            var lay = container.layer;
            for (var i = lay.groupItems.length - 1; i >= 0; i--) {
                var gi = lay.groupItems[i];
                if (gi.name === 'TiledGrid_preview' || gi.name.indexOf('TiledGrid_preview') === 0) {
                    try {
                        gi.remove();
                    } catch (e2) {}
                }
            }
        } catch (e3) {}
    }

    function renderPreview(gapXpt, gapYpt, marginPt, brick, scalePctVal) {
        marginPt = marginPt || 0;
        brick = !!brick;
        var sf = Math.max(0, (typeof scalePctVal === 'number' ? scalePctVal : 100) / 100.0);
        clearPreview();

        // Compute steps and counts using current container/tile bounds
        var tileW = tInfo.width * sf;
        var tileH = tInfo.height * sf;
        var stepXp = tileW + gapXpt;
        var stepYp = tileH + gapYpt;
        if (stepXp <= 0 || stepYp <= 0) return;
        var colsPrev = Math.max(1, Math.ceil((cInfo.width + (brick ? stepXp * 0.5 : 0)) / stepXp));
        var rowsPrev = Math.max(1, Math.ceil(cInfo.height / stepYp));

        var targetLayer = container.layer;
        var group = targetLayer.groupItems.add();
        group.name = 'TiledGrid_preview';
        _previewGroup = group; // set early to ensure we can clear even if interrupted

        var originLeft = cInfo.left;
        var originTop = cInfo.top;

        var adjLeft = cInfo.left + marginPt;
        var adjRight = cInfo.right - marginPt;
        var adjTop = cInfo.top - marginPt;
        var adjBottom = cInfo.bottom + marginPt;

        for (var r = 0; r < rowsPrev; r++) {
            var yTop = originTop - r * stepYp;
            for (var c = 0; c < colsPrev; c++) {
                var xLeft = originLeft + c * stepXp + ((brick && (r % 2 === 1)) ? stepXp * 0.5 : 0);
                var dup = tile.duplicate(group, ElementPlacement.PLACEATBEGINNING);
                try {
                    dup.resize(sf * 100, sf * 100);
                } catch (e) {}
                dup.position = [xLeft, yTop];
            }
        }

        // Trim outside like final
        var containerIsRect = (container.typename === 'PathItem') && TG.isRectanglePath(container);
        var containerIsEllipse = (container.typename === 'PathItem') && TG.isEllipseLike(container);

        var polyPreview = null, polysPreview = null;
        if (!containerIsRect && !containerIsEllipse) {
            if (container.typename === 'PathItem') {
                polysPreview = TG.flattenContainerPolys(container, 12);
                polyPreview = (polysPreview.length ? polysPreview[0] : null);
            } else if (container.typename === 'CompoundPathItem') {
                polysPreview = TG.flattenContainerPolys(container, 12);
            }
        }

        var cxC = (cInfo.left + cInfo.right) / 2.0;
        var cyC = (cInfo.top + cInfo.bottom) / 2.0;
        var rxC = Math.abs(cInfo.right - cInfo.left) / 2.0;
        var ryC = Math.abs(cInfo.top - cInfo.bottom) / 2.0;
        var rxAdj = Math.max(0, rxC - marginPt);
        var ryAdj = Math.max(0, ryC - marginPt);

        function pointInEllipse(x, y) {
            if (rxAdj === 0 || ryAdj === 0) return false;
            var dx = (x - cxC) / rxAdj;
            var dy = (y - cyC) / ryAdj;
            return (dx * dx + dy * dy) <= 1.000001;
        }

        for (var i = group.pageItems.length - 1; i >= 0; i--) {
            var it = group.pageItems[i];
            var ib = TG.boundsInfo(it);
            var keep = false;
            if (containerIsRect) {
                keep = (ib.left >= adjLeft && ib.right <= adjRight && ib.top <= adjTop && ib.bottom >= adjBottom);
            } else if (containerIsEllipse) {
                var corners = [
                    [ib.left, ib.top],
                    [ib.right, ib.top],
                    [ib.left, ib.bottom],
                    [ib.right, ib.bottom]
                ];
                keep = true;
                for (var k = 0; k < 4; k++) {
                    if (!pointInEllipse(corners[k][0], corners[k][1])) {
                        keep = false;
                        break;
                    }
                }
            } else if (container.typename === 'PathItem') {
                if (polysPreview && polysPreview.length) {
                    keep = TG.rectInsidePolys(ib, polysPreview, marginPt);
                } else {
                    var cxTile = (ib.left + ib.right) / 2.0;
                    var cyTile = (ib.top + ib.bottom) / 2.0;
                    keep = (cxTile >= adjLeft && cxTile <= adjRight && cyTile <= adjTop && cyTile >= adjBottom);
                }
            } else if (container.typename === 'CompoundPathItem') {
                if (polysPreview && polysPreview.length) {
                    keep = TG.rectInsidePolys(ib, polysPreview, marginPt);
                } else {
                    var cxTile = (ib.left + ib.right) / 2.0;
                    var cyTile = (ib.top + ib.bottom) / 2.0;
                    keep = (cxTile >= adjLeft && cxTile <= adjRight && cyTile <= adjTop && cyTile >= adjBottom);
                }
            } else {
                var cxTile = (ib.left + ib.right) / 2.0;
                var cyTile = (ib.top + ib.bottom) / 2.0;
                keep = (cxTile >= adjLeft && cxTile <= adjRight && cyTile <= adjTop && cyTile >= adjBottom);
            }
            if (!keep) it.remove();
        }

        try {
            group.opacity = 60;
        } catch (e) {}
        app.redraw();
    }

    // Dialog: spacing between tiles (左右/上下) with Link option
    var gapX = 0,
        gapY = 0,
        marginVal = 0, // margin (pt) — logic TBD
        scalePct = 100, // スケール（%）— ロジックは後で実装
        brickMode = false; // レンガ配置（ロジックは後で実装）
    var _cancel = false;
    (function createSpacingDialog() {
        /* ダイアログボックス / Dialog Window */
        var dlg = new Window('dialog', L('dialogTitle'));
        // Position & Opacity
        var offsetX = 300;
        var dialogOpacity = 0.98;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function() {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);
        dlg.alignChildren = 'fill';

        var unitLabel = TG.getCurrentUnitLabel();
        // Common label width & right align for rows
        var labelWidth = 60; // px

        // Scale row (topmost)
        var rowScale = dlg.add('group');
        rowScale.alignment = ['fill', 'top'];
        rowScale.alignChildren = ['left', 'center'];
        var lblScale = rowScale.add('statictext', undefined, 'スケール');
        try {
            lblScale.justify = 'right';
        } catch (e) {}
        lblScale.preferredSize.width = labelWidth;
        var scaleEdit = rowScale.add('edittext', undefined, '100');
        scaleEdit.characters = 4;
        rowScale.add('statictext', undefined, '（%）');

        changeValueByArrowKey(scaleEdit);
        scaleEdit.onChanging = function() {
            updatePreviewFromFields();
        };

        // Spacing row (panel removed; controls kept)
        var rowS = dlg.add('group');
        rowS.alignment = ['fill', 'top'];
        rowS.alignChildren = ['left', 'center'];
        var lblS = rowS.add('statictext', undefined, '間隔');
        try {
            lblS.justify = 'right';
        } catch (e) {}
        lblS.preferredSize.width = labelWidth;
        var gapEdit = rowS.add('edittext', undefined, '0');
        gapEdit.characters = 4;
        rowS.add('statictext', undefined, '（' + unitLabel + '）');
        gapEdit.active = true;

        // 変更後（allowNegative フラグを追加）
        function changeValueByArrowKey(editText, allowNegative) {
            allowNegative = !!allowNegative;
            editText.addEventListener("keydown", function(event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    if (event.keyName == "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value = Math.floor((value - 1) / delta) * delta;
                        if (!allowNegative && value < 0) value = 0;
                        event.preventDefault();
                    }
                } else if (keyboard.altKey) {
                    delta = 0.1;
                    if (event.keyName == "Up") {
                        value += delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value -= delta;
                        event.preventDefault();
                    }
                } else {
                    delta = 1;
                    if (event.keyName == "Up") {
                        value += delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value -= delta;
                        if (!allowNegative && value < 0) value = 0;
                        event.preventDefault();
                    }
                }

                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                editText.text = value;
                updatePreviewFromFields();
            });
        }

        changeValueByArrowKey(gapEdit);

        function updatePreviewFromFields() {
            var unitCodeNow = TG.getRulerType();
            var s = Math.max(0, parseFloat(gapEdit.text) || 0);
            var m = 0;
            if (typeof marginEdit !== 'undefined') {
                m = parseFloat(marginEdit.text);
                if (isNaN(m)) m = 0;
            }
            s = TG.toPoints(s, unitCodeNow);
            m = TG.toPoints(m, unitCodeNow);
            var b = false;
            try {
                b = !!(brickChk && brickChk.value);
            } catch (e) {}
            var sc = parseFloat(scaleEdit.text);
            if (isNaN(sc)) sc = 100;
            renderPreview(s, s, m, b, sc);
        }
        gapEdit.onChanging = function() {
            updatePreviewFromFields();
        };
        updatePreviewFromFields();

        // Margin row (panel removed; controls kept)
        var rowM = dlg.add('group');
        rowM.alignment = ['fill', 'top'];
        rowM.alignChildren = ['left', 'center'];
        var lblM = rowM.add('statictext', undefined, 'マージン');
        try {
            lblM.justify = 'right';
        } catch (e) {}
        lblM.preferredSize.width = labelWidth;
        var marginEdit = rowM.add('edittext', undefined, '0');
        marginEdit.characters = 4;
        rowM.add('statictext', undefined, '（' + unitLabel + '）');
  
        // Arrow-key increments for margin as well
        changeValueByArrowKey(marginEdit, true); // マージンは負OK


        // Brick layout toggle (grouped & horizontally centered)
        var rowBWrap = dlg.add('group');
        rowBWrap.alignment = ['fill', 'top'];
        rowBWrap.alignChildren = ['center', 'center'];

        var rowB = rowBWrap.add('group');
        rowB.alignChildren = ['left', 'center'];
        var brickChk = rowB.add('checkbox', undefined, 'レンガ');
        brickChk.value = false;
        brickChk.onClick = function() {
            updatePreviewFromFields();
        };


        var btns = dlg.add('group');
        btns.alignment = 'right';
        btns.add('button', undefined, 'Cancel', {
            name: 'cancel'
        });
        btns.add('button', undefined, 'OK', {
            name: 'ok'
        });

        var res = dlg.show();
        if (res !== 1) {
            clearPreview();
            _cancel = true;
            return;
        }

        var unitCode = TG.getRulerType();
        var sVal = Math.max(0, parseFloat(gapEdit.text) || 0);
        scalePct = parseFloat(scaleEdit.text);
        if (isNaN(scalePct)) scalePct = 100;
        if (scalePct < 0) scalePct = 0;
        brickMode = !!(brickChk && brickChk.value);

        marginVal = parseFloat(marginEdit.text);
        if (isNaN(marginVal)) marginVal = 0;
        // Convert from current units to points for geometry
        sVal = TG.toPoints(sVal, unitCode);
        gapX = sVal;
        gapY = sVal;
        marginVal = TG.toPoints(marginVal, unitCode);

    })();
    if (_cancel) {
        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        return;
    }

    clearPreview();

    // Compute steps with spacing and counts to fully cover
    var sf = Math.max(0, (scalePct || 100) / 100.0);
    var tileWFinal = tInfo.width * sf;
    var tileHFinal = tInfo.height * sf;
    var stepX = tileWFinal + gapX;
    var stepY = tileHFinal + gapY;
    var cols = Math.max(1, Math.ceil((cInfo.width + (brickMode ? stepX * 0.5 : 0)) / stepX));
    var rows = Math.max(1, Math.ceil(cInfo.height / stepY));

    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    app.redraw();

    // Create a group for duplicates on the same layer as container
    var targetLayer = container.layer;
    var gridGroup = targetLayer.groupItems.add();
    gridGroup.name = 'TiledGrid';

    // Top-left origin for grid placement
    var originLeft = cInfo.left;
    var originTop = cInfo.top;

    // Duplicate in a grid with specified spacing
    for (var r = 0; r < rows; r++) {
        var yTop = originTop - r * stepY;
        for (var c = 0; c < cols; c++) {
            var xLeft = originLeft + c * stepX + ((brickMode && (r % 2 === 1)) ? stepX * 0.5 : 0);
            var dup = tile.duplicate(gridGroup, ElementPlacement.PLACEATBEGINNING);
            try {
                dup.resize(sf * 100, sf * 100);
            } catch (e) {}
            // Illustrator positions use [left, top]
            dup.position = [xLeft, yTop];
        }
    }

    // --- Mask-less trimming: keep only tiles fully inside the container ---
    var containerIsRect = (container.typename === 'PathItem') && TG.isRectanglePath(container);
    var containerIsEllipse = (container.typename === 'PathItem') && TG.isEllipseLike(container);

    // Precompute ellipse params from visible bounds
    var cx = (cInfo.left + cInfo.right) / 2.0;
    var cy = (cInfo.top + cInfo.bottom) / 2.0;
    var rx = Math.abs(cInfo.right - cInfo.left) / 2.0;
    var ry = Math.abs(cInfo.top - cInfo.bottom) / 2.0;
    var adjLeftF = cInfo.left + marginVal;
    var adjRightF = cInfo.right - marginVal;
    var adjTopF = cInfo.top - marginVal;
    var adjBottomF = cInfo.bottom + marginVal;
    var rxAdjF = Math.max(0, rx - marginVal);
    var ryAdjF = Math.max(0, ry - marginVal);

    var polysFinal = null;
    if (!containerIsRect && !containerIsEllipse) {
        if (container.typename === 'PathItem' || container.typename === 'CompoundPathItem') {
            polysFinal = TG.flattenContainerPolys(container, 16);
        }
    }

    function pointInEllipse(x, y) {
        if (rx === 0 || ry === 0) return false;
        var dx = (x - cx) / rx;
        var dy = (y - cy) / ry;
        return (dx * dx + dy * dy) <= 1.000001; // small tolerance
    }

    // Iterate duplicates and remove those outside the container
    for (var i = gridGroup.pageItems.length - 1; i >= 0; i--) {
        var it = gridGroup.pageItems[i];
        var ib = TG.boundsInfo(it);
        var keep = false;

        if (containerIsRect) {
            // Full tile bbox must fit inside adjusted container bbox
            keep = (ib.left >= adjLeftF && ib.right <= adjRightF && ib.top <= adjTopF && ib.bottom >= adjBottomF);
        } else if (containerIsEllipse) {
            // All four corners must be inside the adjusted ellipse
            var corners = [
                [ib.left, ib.top],
                [ib.right, ib.top],
                [ib.left, ib.bottom],
                [ib.right, ib.bottom]
            ];
            keep = (rxAdjF > 0 && ryAdjF > 0);
            if (keep) {
                for (var k = 0; k < 4; k++) {
                    var dx = (corners[k][0] - cx) / rxAdjF;
                    var dy = (corners[k][1] - cy) / ryAdjF;
                    if ((dx * dx + dy * dy) > 1.000001) {
                        keep = false;
                        break;
                    }
                }
            }
        } else if (container.typename === 'PathItem') {
            if (polysFinal && polysFinal.length) {
                keep = TG.rectInsidePolys(ib, polysFinal, marginVal);
            } else {
                var cxTile = (ib.left + ib.right) / 2.0;
                var cyTile = (ib.top + ib.bottom) / 2.0;
                keep = (cxTile >= adjLeftF && cxTile <= adjRightF && cyTile <= adjTopF && cyTile >= adjBottomF);
            }
        } else if (container.typename === 'CompoundPathItem') {
            if (polysFinal && polysFinal.length) {
                keep = TG.rectInsidePolys(ib, polysFinal, marginVal);
            } else {
                var cxTile = (ib.left + ib.right) / 2.0;
                var cyTile = (ib.top + ib.bottom) / 2.0;
                keep = (cxTile >= adjLeftF && cxTile <= adjRightF && cyTile <= adjTopF && cyTile >= adjBottomF);
            }
        } else {
            var cxTile = (ib.left + ib.right) / 2.0;
            var cyTile = (ib.top + ib.bottom) / 2.0;
            keep = (cxTile >= adjLeftF && cxTile <= adjRightF && cyTile <= adjTopF && cyTile >= adjBottomF);
        }

        if (!keep) it.remove();
    }
    app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
    app.redraw();

    // Finish
}

function __runMain() {
    try {
        main();
    } catch (e) {
        try {
            alert('[TileSmallIntoLarge] Error:\n' + e);
        } catch (_) {}
    }
}
try {
    if (typeof app.doScript === 'function' && typeof ScriptLanguage !== 'undefined' && typeof UndoModes !== 'undefined') {
        app.doScript(__runMain, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'TileSmallIntoLarge');
    } else {
        __runMain();
    }
} catch (e) {
    __runMain();
}