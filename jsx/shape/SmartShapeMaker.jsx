#targetengine "MyScriptEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SmartShapeMaker.jsx (v1.7)

Illustrator script to create custom shapes from a single dialog
(Circle / Polygon / Star / Superellipse / Reuleaux-style).
Real-time preview, adjustable sides, width, rotation, and advanced options are supported.
Japanese / English UI.

### 更新日 / Updated:
- 20260228

Main Features:
- Specify number of sides (0 = Circle, 3/4/5/6/8, or custom with slider)
- Circle panel:
  - Superellipse option (only when sides = 0)
  - Superellipse shape control (exponent)
  - Anchor Points panel (2 / 3 / 4 / 5 / 6)
  - When Superellipse is ON:
    - Rotate is forced OFF
    - Live Shape is forced OFF
    - Anchor Points panel is dimmed
- Star panel:
  - Star option + Pentagram option (side-by-side)
  - Inner radius input + 0–100 slider
  - Inner radius controls are dimmed when Star is OFF
  - When Pentagram is ON, Rotate is forced OFF
- Triangle direction options (Left / Right / Down) when sides = 3
- Width (size) panel with unit display
- Rotation panel:
  - Auto angle is used when Rotate is OFF (Circle=45°, Polygon=360/(sides*2))
  - When sides = 3 and Rotate is enabled, Triangle direction defaults to “Down” (60°)
  - Arrow-key editing supported
- Reuleaux-style option (odd-sided polygons only)
  - Adjustable appearance amount (0–200%)
  - Amount resets to 100% when Reuleaux is enabled
- Options panel:
  - Live Shape conversion (Convert to Shape) on finalize
  - Split at Anchor Points (creates open stroked segments)
- Dialog opacity and position are restored within the current Illustrator session
- Preview does not pollute Undo history; final result can be undone in a single step
- View Zoom slider above OK / Cancel

Keyboard Shortcuts:
- E : Circle (0)
- A : Toggle Rotate
- S : Toggle Star
- P : Toggle Pentagram
- L : Triangle Left (also sets sides = 3)
- R : Triangle Right (also sets sides = 3)
- B : Triangle Down (also sets sides = 3)
- D : Toggle Split at Anchor Points

Usage Flow:
1. Set sides, width, star/circle options, rotation, and options in the dialog
2. Preview updates in real-time
3. Click OK to finalize the preview object at the artboard center

Original Idea: Seiji Miyazawa (Sankai Lab)
https://x.com/onthehead/status/2007350198721483172

Version: v1.7 (20260228)
*/

// Language detection
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var SCRIPT_VERSION = "v1.7";

var LABELS = {
    dialogTitle: {
        ja: "基本図形の作成 " + SCRIPT_VERSION,
        en: "Create Basic Shapes " + SCRIPT_VERSION
    },
    shapeType: { ja: "辺の数", en: "Sides" },
    circle: { ja: "円", en: "Circle" },
    custom: { ja: "それ以外", en: "Other" },
    starPanel: { ja: "スター", en: "Star" },
    circlePanel: { ja: "円", en: "Circle" },
    anchorPanel: { ja: "アンカーポイント", en: "Anchor Points" },
    superEllipse: { ja: "スーパー楕円", en: "Superellipse" },
    trianglePanel: { ja: "三角形", en: "Triangle" },
    triangleRight: { ja: "右", en: "Right" },
    triangleLeft: { ja: "左", en: "Left" },
    triangleDown: { ja: "下", en: "Down" },
    star: { ja: "スター", en: "Star" },
    innerRadius: { ja: "第2半径：", en: "Inner Radius:" },
    percent: { ja: "%", en: "%" },
    pentagram: { ja: "五芒星", en: "Pentagram" },
    rotation: { ja: "回転", en: "Rotate" },
    width: { ja: "幅：", en: "Width:" },
    widthPanel: { ja: "幅", en: "Width" },
    optionPanel: { ja: "オプション", en: "Options" },
    reuleaux: { ja: "ルーロー（定幅図形）", en: "Reuleaux (Constant Width)" },
    splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at Anchor Points" },
    liveShape: { ja: "ライブシェイプ化", en: "Live Shape" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    viewZoom: { ja: "画面ズーム", en: "View Zoom" },
    previewError: { ja: "プレビューエラー：", en: "Preview Error: " },
    finalError: { ja: "確定エラー：", en: "Final Error: " }
};

var previewShape = null;
var applyLiveShape = true;

// Session-only dialog state (kept only while Illustrator is running)
// Uses the persistent engine specified by #targetengine
var SESSION_STATE_KEY = "__SmartShapeMaker_State__";
if (!$.global[SESSION_STATE_KEY]) {
    $.global[SESSION_STATE_KEY] = {};
}
function getSessionState() {
    return $.global[SESSION_STATE_KEY];
}

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var unit = getRulerUnitInfo();
    var result = showInputDialog(unit.label, unit.factor);
    if (!result) return;

    // Re-acquire activeDocument in case user changed documents while the dialog was open
    doc = app.activeDocument;

    finalizeShape(doc);
    if (applyLiveShape) app.executeMenuCommand('Convert to Shape');
}

// Preview / Undo manager
// - rollback(): undo all preview operations
// - confirm(finalAction): rollback previews, then apply the final action once
function PreviewManager() {
    this.undoDepth = 0;

    this.addStep = function (func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            alert(LABELS.previewError[lang] + e);
        }
    };

    this.rollback = function () {
        while (this.undoDepth > 0) {
            try { app.undo(); } catch (e) { break; }
            this.undoDepth--;
        }
        try { app.redraw(); } catch (e) { }
    };

    this.confirm = function (finalAction) {
        if (finalAction) {
            this.rollback();
            try { finalAction(); } catch (e) { alert(LABELS.finalError[lang] + e); }
            this.undoDepth = 0;
        } else {
            this.undoDepth = 0;
        }
    };
}

// Adjust value with arrow keys (Up/Down)
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;
        if (event.keyName == "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName == "Down") {
            value -= delta;
            event.preventDefault();
        }
        editText.text = value;

        // If an onChanging handler exists (used for preview), invoke it so arrow-key edits update the preview.
        if (typeof editText.onChanging === "function") {
            try { editText.onChanging(); } catch (e) { }
        }
    });
}

// Get ruler unit label and factor for conversion
function getRulerUnitInfo() {
    var t = app.preferences.getIntegerPreference("rulerType");
    var u = { label: "pt", factor: 1.0 };
    if (t === 0) u = { label: "inch", factor: 72.0 };
    else if (t === 1) u = { label: "mm", factor: 72.0 / 25.4 };
    else if (t === 3) u = { label: "pica", factor: 12.0 };
    else if (t === 4) u = { label: "cm", factor: 72.0 / 2.54 };
    else if (t === 5) u = { label: "Q", factor: 72.0 / 25.4 * 0.25 };
    else if (t === 6) u = { label: "px", factor: 1.0 };
    return u;
}

// Format angle display with up to 3 decimals or integer
function formatAngle(value) {
    var rounded = Math.round(value * 1000) / 1000;
    return (rounded % 1 === 0) ? String(Math.round(rounded)) : String(rounded);
}

// Math helper
function sign(x) {
    return ((x > 0) - (x < 0)) || +x;
}

// Create a Superellipse path (rounded-rectangle like circle variant)
// Based on the reference superellipse script: sample points + smoothing handles.
function createSuperellipsePath(doc, sizePt, exponent, numPoints) {
    exponent = (typeof exponent === 'number' && exponent > 0) ? exponent : 2.5;
    numPoints = (typeof numPoints === 'number' && numPoints >= 8) ? Math.round(numPoints) : 8;

    var layer = doc.activeLayer;
    layer.locked = false;
    layer.visible = true;

    var center = doc.activeView.centerPoint;
    var cx = center[0];
    var cy = center[1];

    var w = sizePt;
    var h = sizePt;

    var TWO_PI = Math.PI * 2;
    var anchors = [];
    for (var i = 0; i < numPoints; i++) {
        var theta = (TWO_PI * i) / numPoints;
        var cosT = Math.cos(theta);
        var sinT = Math.sin(theta);
        var x = Math.pow(Math.abs(cosT), 2 / exponent) * (w / 2) * sign(cosT);
        var y = Math.pow(Math.abs(sinT), 2 / exponent) * (h / 2) * sign(sinT);
        anchors.push([x + cx, y + cy]);
    }

    var pathItem = layer.pathItems.add();
    pathItem.setEntirePath(anchors);
    pathItem.closed = true;

    // Smooth handles
    try {
        var pts = pathItem.pathPoints;
        var n = pts.length;
        if (n >= 4) {
            for (var k = 0; k < n; k++) {
                var prev = anchors[(k - 1 + n) % n];
                var cur = anchors[k];
                var next = anchors[(k + 1) % n];

                // Tangent vector (next - prev)
                var tx = next[0] - prev[0];
                var ty = next[1] - prev[1];
                var tlen = Math.sqrt(tx * tx + ty * ty);
                if (tlen === 0) continue;
                tx /= tlen;
                ty /= tlen;

                // Segment lengths
                var d1x = cur[0] - prev[0];
                var d1y = cur[1] - prev[1];
                var d2x = next[0] - cur[0];
                var d2y = next[1] - cur[1];
                var d1 = Math.sqrt(d1x * d1x + d1y * d1y);
                var d2 = Math.sqrt(d2x * d2x + d2y * d2y);

                var hlen = Math.min(d1, d2) * 0.35;
                var left = [cur[0] - tx * hlen, cur[1] - ty * hlen];
                var right = [cur[0] + tx * hlen, cur[1] + ty * hlen];

                pts[k].anchor = cur;
                pts[k].leftDirection = left;
                pts[k].rightDirection = right;
                pts[k].pointType = PointType.SMOOTH;
            }
        }
    } catch (e) {
        // ignore
    }

    // Appearance (match existing createShape defaults)
    pathItem.filled = true;
    pathItem.fillColor = doc.defaultFillColor;
    pathItem.stroked = false;

    return pathItem;
}

// Create a smooth circle-like closed path with N anchors (N>=3).
// Uses cubic-bezier handle length k = 4/3 * tan(pi/(2N)).
function createCirclePathWithNAnchors(doc, sizePt, N) {
    var layer = doc.activeLayer;
    layer.locked = false;
    layer.visible = true;

    var center = doc.activeView.centerPoint;
    var cx = center[0];
    var cy = center[1];

    var r = sizePt / 2;
    N = (typeof N === 'number') ? Math.round(N) : 4;
    if (N < 2) N = 2;

    var k = (4 / 3) * Math.tan(Math.PI / (2 * N));
    var hlen = r * k;

    // Start at -90° so one anchor is at the top.
    var anchors = [];
    for (var i = 0; i < N; i++) {
        var a = (-Math.PI / 2) + (2 * Math.PI * i) / N;
        anchors.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
    }

    var pathItem = layer.pathItems.add();
    pathItem.setEntirePath(anchors);
    pathItem.closed = true;

    // Smooth handles (tangents)
    try {
        var pts = pathItem.pathPoints;
        for (var j = 0; j < N; j++) {
            var ax = anchors[j][0];
            var ay = anchors[j][1];
            var ang = (-Math.PI / 2) + (2 * Math.PI * j) / N;

            // Tangent direction is perpendicular to radius: [-sin, cos]
            var tx = -Math.sin(ang);
            var ty = Math.cos(ang);

            pts[j].anchor = [ax, ay];
            pts[j].leftDirection = [ax - tx * hlen, ay - ty * hlen];
            pts[j].rightDirection = [ax + tx * hlen, ay + ty * hlen];
            pts[j].pointType = PointType.SMOOTH;
        }
    } catch (e) {
        // ignore
    }

    pathItem.filled = true;
    pathItem.fillColor = doc.defaultFillColor;
    pathItem.stroked = false;

    return pathItem;
}

// Convert an odd-sided polygon into a Reuleaux (constant-width) polygon by turning each edge into a circular arc.
// This adapts the logic from the reference script `reuleaux_polygon.jsx`.
function applyReuleauxToPolygon(pathItem, amount) {
    try {
        if (!pathItem || pathItem.typename !== "PathItem") return pathItem;
        if (!pathItem.pathPoints || pathItem.pathPoints.length < 3) return pathItem;
        var pts = pathItem.pathPoints;
        var numPts = pts.length;
        if (numPts % 2 === 0) return pathItem; // odd only

        // amount: 0.0 .. 2.0 (1.0 = current)
        amount = (typeof amount === "number") ? amount : 1;
        if (isNaN(amount)) amount = 1;
        if (amount < 0) amount = 0;
        if (amount > 2) amount = 2;

        // Cache anchor coordinates
        var p = [];
        for (var i = 0; i < numPts; i++) {
            p.push([pts[i].anchor[0], pts[i].anchor[1]]);
        }

        for (var i = 0; i < numPts; i++) {
            var idx1 = i;
            var idx2 = (i + 1) % numPts;
            var idxCenter = (i + Math.floor((numPts + 1) / 2)) % numPts;

            var P1 = p[idx1];
            var P2 = p[idx2];
            var P0 = p[idxCenter];

            var V1 = [P1[0] - P0[0], P1[1] - P0[1]];
            var V2 = [P2[0] - P0[0], P2[1] - P0[1]];

            var Z = V1[0] * V2[1] - V1[1] * V2[0];

            var R1 = Math.sqrt(V1[0] * V1[0] + V1[1] * V1[1]);
            var R2 = Math.sqrt(V2[0] * V2[0] + V2[1] * V2[1]);
            if (R1 === 0 || R2 === 0) continue;
            var R = (R1 + R2) / 2;

            var dot = V1[0] * V2[0] + V1[1] * V2[1];
            var cosTheta = dot / (R1 * R2);
            if (cosTheta < -1) cosTheta = -1;
            if (cosTheta > 1) cosTheta = 1;
            var deltaTheta = Math.acos(cosTheta);

            var L = R * (4 / 3) * Math.tan(deltaTheta / 4);
            L *= amount; // ★度合い

            var T1, T2;
            if (Z > 0) {
                T1 = [-V1[1], V1[0]];
                T2 = [V2[1], -V2[0]];
            } else {
                T1 = [V1[1], -V1[0]];
                T2 = [-V2[1], V2[0]];
            }

            var lenT1 = Math.sqrt(T1[0] * T1[0] + T1[1] * T1[1]);
            var lenT2 = Math.sqrt(T2[0] * T2[0] + T2[1] * T2[1]);
            if (lenT1 === 0 || lenT2 === 0) continue;

            var rightDir = [P1[0] + L * T1[0] / lenT1, P1[1] + L * T1[1] / lenT1];
            var leftDir = [P2[0] + L * T2[0] / lenT2, P2[1] + L * T2[1] / lenT2];

            pts[idx1].rightDirection = rightDir;
            pts[idx2].leftDirection = leftDir;

            pts[idx1].pointType = PointType.CORNER;
            pts[idx2].pointType = PointType.CORNER;
        }

        pathItem.closed = true;
    } catch (e) { }
    return pathItem;
}

// Get selected side count from radio buttons or custom input
function getSelectedSideValue(radios, input) {
    for (var i = 0; i < radios.length; i++) {
        if (radios[i].value) {
            return (i === 6) ? parseInt(input.text, 10) : [0, 3, 4, 5, 6, 8][i];
        }
    }
    return 4;
}

// Finalize shape selection and clear preview reference
function finalizeShape(doc) {
    if (!previewShape) return;
    doc.selection = [previewShape];
    previewShape = null;
}

// Create shape based on parameters
function createShape(doc, sizePt, sides, isStar, innerRatio, rotateEnabled, rotateAngle, splitAtAnchors, useSuperEllipse, superExp, circleAnchorCount, useReuleaux, reuleauxAmount) {
    var layer = doc.activeLayer;
    layer.locked = false;
    layer.visible = true;
    var center = doc.activeView.centerPoint;
    var radius = sizePt / 2;
    var innerRadius = radius * (innerRatio / 100);
    var shape;

    if (sides === 0) {
        if (useSuperEllipse) {
            shape = createSuperellipsePath(doc, sizePt, superExp);
        } else {
            // Default 4 anchors uses Illustrator ellipse; other counts create a custom smooth path.
            var nAnchors = (typeof circleAnchorCount === 'number') ? Math.round(circleAnchorCount) : 4;
            if (nAnchors < 2) nAnchors = 2;
            if (nAnchors === 4) {
                shape = layer.pathItems.ellipse(center[1] + radius, center[0] - radius, sizePt, sizePt);
            } else {
                shape = createCirclePathWithNAnchors(doc, sizePt, nAnchors);
            }
        }
    } else if (isStar) {
        shape = doc.pathItems.star(center[0], center[1], radius, innerRadius, sides);
    } else {
        shape = doc.pathItems.polygon(center[1], center[0], radius, sides);
    }

    shape.filled = true;
    shape.fillColor = doc.defaultFillColor;
    shape.stroked = false;

    var b = shape.geometricBounds;
    var cx = (b[0] + b[2]) / 2;
    var cy = (b[1] + b[3]) / 2;
    shape.translate(center[0] - cx, center[1] - cy);

    // Reuleaux (constant-width) conversion for odd-sided polygons
    if (useReuleaux && !isStar && sides > 0 && (sides % 2 === 1)) {
        try {
            // `polygon(...)` returns a PathItem; convert its edges into arcs.
            shape = applyReuleauxToPolygon(shape, reuleauxAmount);
        } catch (e) { }
    }

    if (rotateEnabled && !isNaN(rotateAngle)) {
        shape.rotate(rotateAngle, true, true, true, true, Transformation.CENTER);
    }
    if (splitAtAnchors) {
        shape = splitPathAtAnchors(doc, shape);
    }
    doc.selection = [shape];
    return shape;
}

// Split a path into segments at each anchor point.
// Returns a GroupItem containing open PathItems (one per segment).
function splitPathAtAnchors(doc, pathItem) {
    if (!pathItem || !pathItem.pathPoints || pathItem.pathPoints.length < 2) return pathItem;

    var layer = doc.activeLayer;
    var group = layer.groupItems.add();

    var pts = pathItem.pathPoints;
    var n = pts.length;
    var closed = pathItem.closed;

    for (var i = 0; i < n; i++) {
        var j = i + 1;
        if (j >= n) {
            if (!closed) break;
            j = 0;
        }

        var p0 = pts[i];
        var p1 = pts[j];

        // Create an open path for this segment
        var seg = group.pathItems.add();
        seg.closed = false;

        // Set anchors first
        seg.setEntirePath([p0.anchor, p1.anchor]);

        // Copy bezier handles for the segment
        // For an open 2-point path: p0 uses rightDirection, p1 uses leftDirection.
        seg.pathPoints[0].leftDirection = p0.anchor;
        seg.pathPoints[0].rightDirection = p0.rightDirection;
        seg.pathPoints[0].pointType = p0.pointType;

        seg.pathPoints[1].leftDirection = p1.leftDirection;
        seg.pathPoints[1].rightDirection = p1.anchor;
        seg.pathPoints[1].pointType = p1.pointType;

        // Appearance: segments should be stroked (fill doesn't make sense for open paths)
        seg.filled = false;
        seg.stroked = true;

        // Stroke: black, 0.3pt (RGB / CMYK supported)
        try {
            var color;
            if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) {
                color = new CMYKColor();
                color.cyan = 0; color.magenta = 0; color.yellow = 0; color.black = 100;
            } else {
                color = new RGBColor();
                color.red = 0; color.green = 0; color.blue = 0;
            }
            seg.strokeColor = color;
        } catch (e) { }
        try { seg.strokeWidth = 0.3; } catch (e) { }
    }

    // Remove original path
    try { pathItem.remove(); } catch (e) { }

    return group;
}

// Show input dialog and handle UI and events
function showInputDialog(unitLabel, unitFactor) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    var previewMgr = new PreviewManager();
    var doc = app.activeDocument;
    // Dialog appearance (opacity) and initial position offset
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(targetDlg, dx, dy) {
        try {
            var currentX = targetDlg.location[0];
            var currentY = targetDlg.location[1];
            targetDlg.location = [currentX + dx, currentY + dy];
        } catch (e) { }
    }

    function getSavedDialogLocation() {
        try {
            var st = getSessionState();
            if (st && st.dlgLocation && st.dlgLocation.length === 2) {
                var x = Number(st.dlgLocation[0]);
                var y = Number(st.dlgLocation[1]);
                if (!isNaN(x) && !isNaN(y)) return [x, y];
            }
        } catch (e) { }
        return null;
    }

    function saveDialogLocation(targetDlg) {
        try {
            if (!targetDlg || !targetDlg.location) return;
            var st = getSessionState();
            st.dlgLocation = [Number(targetDlg.location[0]), Number(targetDlg.location[1])];
        } catch (e) { }
    }

    function getCircleAnchorCountFromRadios() {
        try {
            return circleAnchorRadios.r2.value ? 2 :
                (circleAnchorRadios.r3.value ? 3 :
                    (circleAnchorRadios.r5.value ? 5 :
                        (circleAnchorRadios.r6.value ? 6 : 4)));
        } catch (e) {
            return 4;
        }
    }

    function setDialogOpacity(targetDlg, opacityValue) {
        try {
            targetDlg.opacity = opacityValue;
        } catch (e) { }
    }
    // Keyboard shortcuts
    dlg.addEventListener("keydown", function (e) {
        if (!e || !e.keyName) return;

        switch (e.keyName.toUpperCase()) {

            case "E":
                // 0 sides (Circle)
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[0].value = true;
                customInput.enabled = false;
                applyAutoRotationForSides(0);
                updatePreview();
                e.preventDefault();
                break;

            // Triangle directions and split toggle (inserted here)
            case "L":
                // Triangle Left
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
                applyAutoRotationForSides(3);
                triangleLeftRadio.value = true;
                onTriangleDirectionChange();
                e.preventDefault();
                break;

            case "R":
                // Triangle Right
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
                applyAutoRotationForSides(3);
                triangleRightRadio.value = true;
                onTriangleDirectionChange();
                e.preventDefault();
                break;

            case "B":
                // Triangle Down (Bottom)
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
                applyAutoRotationForSides(3);
                triangleDownRadio.value = true;
                onTriangleDirectionChange();
                e.preventDefault();
                break;

            case "D":
                // Toggle Split at Anchor Points
                splitAtAnchorsCheck.value = !splitAtAnchorsCheck.value;
                if (typeof splitAtAnchorsCheck.onClick === "function") {
                    splitAtAnchorsCheck.onClick();
                } else {
                    updatePreview();
                }
                e.preventDefault();
                break;

            case "A":
                // Rotate toggle
                rotateCheck.value = !rotateCheck.value;
                rotateInput.enabled = rotateCheck.value;
                rotateLabel.enabled = rotateCheck.value;
                if (rotateCheck.value) {
                    applyDefaultRotationWhenEnablingRotate();
                }
                updatePreview();
                e.preventDefault();
                break;

            case "S":
                // Star toggle
                starCheck.value = !starCheck.value;
                if (!starCheck.value) {
                    pentagramCheck.value = false;
                }
                updatePreview();
                e.preventDefault();
                break;

            case "P":
                // Pentagram toggle (only when Star is enabled)
                if (!starCheck.value) {
                    starCheck.value = true;
                }
                pentagramCheck.value = !pentagramCheck.value;
                updatePreview();
                e.preventDefault();
                break;
        }
    });
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var radios = [], customInput;
    var main = dlg.add("group");
    main.orientation = "row";

    // Left column container (panel + rotation row)
    var leftCol = main.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";
    leftCol.alignment = "top";

    var left = leftCol.add("panel", undefined, LABELS.shapeType[lang]);
    left.orientation = "column";
    left.alignChildren = "left";
    left.margins = [20, 20, 10, 10];

    radios[0] = left.add("radiobutton", undefined, "0 (" + LABELS.circle[lang] + ")");
    radios[1] = left.add("radiobutton", undefined, "3");
    radios[2] = left.add("radiobutton", undefined, "4");
    radios[3] = left.add("radiobutton", undefined, "5");
    radios[4] = left.add("radiobutton", undefined, "6");
    radios[5] = left.add("radiobutton", undefined, "8");

    var customGroup = left.add("group");
    radios[6] = customGroup.add("radiobutton", undefined, LABELS.custom[lang]);
    customInput = customGroup.add("edittext", undefined, "12");
    customInput.characters = 3;
    customInput.enabled = false;
    changeValueByArrowKey(customInput);

    // Custom sides slider (3–36)
    var customSliderRow = left.add("group");
    customSliderRow.orientation = "row";
    customSliderRow.alignChildren = ["left", "center"];

    var customSlider = customSliderRow.add("slider", undefined, 12, 3, 36);
    customSlider.preferredSize.width = 180;
    customSlider.enabled = false;
    radios[2].value = true;

    // Rotation panel placed under the sides panel (left column)
    var rotatePanel = leftCol.add("panel", undefined, LABELS.rotation[lang]);
    rotatePanel.orientation = "column";
    rotatePanel.alignChildren = "left";
    rotatePanel.margins = [15, 20, 15, 10];

    var rotateRow = rotatePanel.add("group");
    rotateRow.orientation = "row";
    rotateRow.alignChildren = ["left", "center"];
    rotateRow.spacing = 6;

    var rotateCheck = rotateRow.add("checkbox", undefined, "");

    var rotateInput = rotateRow.add("edittext", undefined, "90");
    rotateInput.characters = 4;
    changeValueByArrowKey(rotateInput);

    var rotateLabel = rotateRow.add("statictext", undefined, "°");
    // Initial state: manual rotation only when checked
    rotateInput.enabled = rotateCheck.value;
    rotateLabel.enabled = rotateCheck.value;

    // Width panel placed under the rotation row (left column)
    var widthPanel = leftCol.add("panel", undefined, LABELS.widthPanel[lang]);
    widthPanel.orientation = "column";
    widthPanel.alignChildren = "left";
    widthPanel.margins = [15, 20, 15, 10];

    // Width input moved from bottom area into this panel
    var widthRow = widthPanel.add("group");
    widthRow.orientation = "row";
    widthRow.alignChildren = ["left", "center"];

    // widthRow.add("statictext", undefined, LABELS.width[lang]);
    var sizeInput = widthRow.add("edittext", undefined, "100");
    sizeInput.characters = 5;
    changeValueByArrowKey(sizeInput);
    widthRow.add("statictext", undefined, "(" + unitLabel + ")");

    // Right column container
    var right = main.add("group");
    right.orientation = "column";
    right.alignChildren = "fill";
    right.alignment = "top";

    // Star panel
    var starPanel = right.add("panel", undefined, LABELS.starPanel[lang]);
    starPanel.orientation = "column";
    starPanel.alignChildren = "left";
    starPanel.margins = [20, 20, 10, 10];

    // Create the row group for Star and Pentagram checkboxes
    var starRow = starPanel.add("group");
    starRow.orientation = "row";
    starRow.alignChildren = ["left", "center"];
    starRow.spacing = 10;

    var starCheck = starRow.add("checkbox", undefined, LABELS.star[lang]);
    var pentagramCheck = starRow.add("checkbox", undefined, LABELS.pentagram[lang]);
    pentagramCheck.value = false;

    var innerGroup = starPanel.add("group");
    var innerRadiusLabel = innerGroup.add("statictext", undefined, LABELS.innerRadius[lang]);
    var innerRatioInput = innerGroup.add("edittext", undefined, "30");
    innerRatioInput.characters = 4;
    changeValueByArrowKey(innerRatioInput);
    var innerPercentLabel = innerGroup.add("statictext", undefined, LABELS.percent[lang]);

    // Inner radius slider (0–100)
    var innerSliderRow = starPanel.add("group");
    innerSliderRow.orientation = "row";
    innerSliderRow.alignChildren = ["left", "center"];

    var innerRatioSlider = innerSliderRow.add("slider", undefined, 30, 0, 100);
    innerRatioSlider.preferredSize.width = 200;

    // Circle options panel placed under the Star panel
    var circlePanel = right.add("panel", undefined, LABELS.circlePanel[lang]);
    circlePanel.orientation = "column";
    circlePanel.alignChildren = "left";
    circlePanel.margins = [15, 20, 15, 10];

    var superEllipseCheck = circlePanel.add("checkbox", undefined, LABELS.superEllipse[lang]);
    superEllipseCheck.value = false;

    // Superellipse shape (exponent)
    var superExpRow = circlePanel.add("group");
    superExpRow.orientation = "row";
    superExpRow.alignChildren = ["left", "center"];
    superExpRow.spacing = 8;

    // var superExpLabel = superExpRow.add("statictext", undefined, LABELS.superEllipseExponent[lang]);
    var superExpInput = superExpRow.add("edittext", undefined, "2.5");
    superExpInput.characters = 4;
    changeValueByArrowKey(superExpInput);

    var superExpSlider = superExpRow.add("slider", undefined, 2.5, 1.5, 6.0);
    superExpSlider.preferredSize.width = 150;

    // Anchor points panel placed under the Circle panel
    var anchorPanel = circlePanel.add("panel", undefined, LABELS.anchorPanel[lang]);
    anchorPanel.orientation = "column";
    anchorPanel.alignChildren = "left";
    anchorPanel.margins = [15, 20, 15, 10];

    var anchorRow = anchorPanel.add("group");
    anchorRow.orientation = "column";
    anchorRow.alignChildren = "left";
    anchorRow.spacing = 10;

    var anchorRadioRow = anchorRow.add("group");
    anchorRadioRow.orientation = "row";
    anchorRadioRow.alignChildren = ["left", "center"];
    anchorRadioRow.spacing = 10;

    var circleAnchorRadios = {};
    circleAnchorRadios.r2 = anchorRadioRow.add("radiobutton", undefined, "2");
    circleAnchorRadios.r3 = anchorRadioRow.add("radiobutton", undefined, "3");
    circleAnchorRadios.r4 = anchorRadioRow.add("radiobutton", undefined, "4");
    circleAnchorRadios.r5 = anchorRadioRow.add("radiobutton", undefined, "5");
    circleAnchorRadios.r6 = anchorRadioRow.add("radiobutton", undefined, "6");
    circleAnchorRadios.r4.value = true; // default

    function updateCirclePanelEnabled(sidesValue) {
        var enable = (sidesValue === 0);
        circlePanel.enabled = enable;
        try { anchorPanel.enabled = enable; } catch (e) { }
        if (!enable) {
            try { superEllipseCheck.value = false; } catch (e) { }
            try {
                circleAnchorRadios.r2.value = false;
                circleAnchorRadios.r3.value = false;
                circleAnchorRadios.r4.value = true;
                circleAnchorRadios.r5.value = false;
                circleAnchorRadios.r6.value = false;
                // also reset dim state
                circleAnchorRadios.r2.enabled = true;
                circleAnchorRadios.r3.enabled = true;
                circleAnchorRadios.r4.enabled = true;
                circleAnchorRadios.r5.enabled = true;
                circleAnchorRadios.r6.enabled = true;
                // anchorCountLabel.enabled = true;
            } catch (e) { }
        }
    }

    function updateStarPanelEnabled(sidesValue) {
        // Star options are not applicable for Circle (sides=0)
        var enableStar = (sidesValue !== 0);
        starPanel.enabled = enableStar;
        if (!enableStar) {
            // Clear star-related options when the panel is disabled
            starCheck.value = false;
            pentagramCheck.value = false;
            pentagramCheck.enabled = false;
        }
    }

    function updateInnerRadiusEnabled() {
        // Enabled only when Star is ON
        var en = !!starCheck.value;
        try {
            innerRadiusLabel.enabled = en;
            innerRatioInput.enabled = en;
            innerPercentLabel.enabled = en;
            innerRatioSlider.enabled = en;
        } catch (e) { }
    }

    function updateReuleauxAvailability(sidesValue) {
        try {
            // If Star is ON, skip even/odd rule entirely (Star logic controls state)
            if (starCheck && starCheck.value) {
                // Star logic controls Reuleaux; keep amount UI in sync
                try { updateReuleauxAmountEnabled(); } catch (e) { }
                return;
            }

            // Reuleaux is available only for odd-numbered polygons (3,5,7,...) and not for Circle (0) or even sides.
            var enable = (typeof sidesValue === "number") && (sidesValue > 0) && (sidesValue % 2 === 1);
            reuleauxCheck.enabled = enable;
            if (!enable) reuleauxCheck.value = false;
            updateReuleauxAmountEnabled();
        } catch (e) { }
    }

    function clampSuperExp(v) {
        v = Number(v);
        if (isNaN(v)) v = 2.5;
        if (v < 1.5) v = 1.5;
        if (v > 6.0) v = 6.0;
        // keep 1 decimal
        v = Math.round(v * 10) / 10;
        return v;
    }

    function syncSuperExpUI(v) {
        v = clampSuperExp(v);
        try {
            superExpInput.text = String(v);
            superExpSlider.value = v;
        } catch (e) { }
        return v;
    }

    function clampReuleauxAmount(v) {
        v = Math.round(Number(v));
        if (isNaN(v)) v = 100;
        if (v < 0) v = 0;
        if (v > 200) v = 200;
        return v;
    }

    function syncReuleauxAmountUI(v) {
        v = clampReuleauxAmount(v);
        try {
            reuleauxAmountInput.text = String(v);
            reuleauxAmountSlider.value = v;
        } catch (e) { }
        return v;
    }

    function updateReuleauxAmountEnabled() {
        try {
            var en = (reuleauxCheck.enabled && reuleauxCheck.value);
            reuleauxAmountInput.enabled = en;
            reuleauxAmountSlider.enabled = en;
        } catch (e) { }
    }

    function updateSuperellipseSmoothnessEnabled(sidesValue) {
        var superEff = (superEllipseCheck.value && sidesValue === 0);
        try {
            // Exponent UI
            superExpInput.enabled = superEff;
            superExpSlider.enabled = superEff;

            // Anchor panel should be disabled when Superellipse is ON
            anchorPanel.enabled = !superEff;
        } catch (e) { }
    }


    // Triangle panel moved into the Rotation panel
    // (Spacer removed)
    var trianglePanel = rotatePanel.add("panel", undefined, LABELS.trianglePanel[lang]);
    trianglePanel.orientation = "column";
    trianglePanel.alignChildren = "left";
    trianglePanel.margins = [15, 20, 15, 10];

    // Options panel placed under the Triangle panel
    var optionPanel = right.add("panel", undefined, LABELS.optionPanel[lang]);
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.margins = [15, 20, 15, 10];

    // Live shape option
    var liveShapeCheck = optionPanel.add("checkbox", undefined, LABELS.liveShape[lang]);
    liveShapeCheck.value = true;

    // Split at anchor points
    var splitAtAnchorsCheck = optionPanel.add("checkbox", undefined, LABELS.splitAtAnchors[lang]);
    splitAtAnchorsCheck.value = false;

    // Reuleaux (constant-width shape) (logic TBD)
    var reuleauxCheck = optionPanel.add("checkbox", undefined, LABELS.reuleaux[lang]);
    reuleauxCheck.value = false;

    // Reuleaux amount (appearance). 100 = default, 0..200
    var reuleauxAmountRow = optionPanel.add("group");
    reuleauxAmountRow.orientation = "row";
    reuleauxAmountRow.alignChildren = ["left", "center"];
    reuleauxAmountRow.spacing = 8;

    // (Label omitted for compact UI)
    var reuleauxAmountInput = reuleauxAmountRow.add("edittext", undefined, "100");
    reuleauxAmountInput.characters = 4;
    changeValueByArrowKey(reuleauxAmountInput);

    var reuleauxAmountSlider = reuleauxAmountRow.add("slider", undefined, 100, 0, 200);
    reuleauxAmountSlider.preferredSize.width = 150;

    function updateLiveShapeAvailability(isSplit, isSuperEllipseEffective, isCustomCircleAnchors, isReuleaux) {
        if (isSplit || isSuperEllipseEffective || isCustomCircleAnchors || isReuleaux) {
            liveShapeCheck.value = false;
            liveShapeCheck.enabled = false;
        } else {
            liveShapeCheck.enabled = true;
        }
    }

    function refreshLiveShapeAvailabilityFromUI() {
        try {
            var sidesNow = getSelectedSideValue(radios, customInput);
            var superEff = (superEllipseCheck.value && sidesNow === 0);
            var n = getCircleAnchorCountFromRadios();
            if (!n || n < 2) n = 2;
            // Live Shape can be enabled only when Circle anchors == 4 (and not Superellipse)
            var isCustomCircleAnchors = (sidesNow === 0 && !superEff && n !== 4);
            updateLiveShapeAvailability(splitAtAnchorsCheck.value, superEff, isCustomCircleAnchors, reuleauxCheck.value);
        } catch (e) {
            // Conservative fallback
            try {
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            } catch (_) { }
        }
    }

    // Update rotation input to the "auto" value used when Rotate is OFF.
    // This is used when the number of sides changes, even if Rotate is ON.
    function applyAutoRotationForSides(sidesValue) {
        var angle;
        if (sidesValue === 0) {
            angle = 45;
        } else if (sidesValue >= 3) {
            angle = 360 / (sidesValue * 2);
        } else {
            return;
        }
        rotateInput.text = formatAngle(angle);
    }

    // When enabling Rotate, pick a sensible default angle.
    // Circle (sides=0) custom anchor count presets when enabling Rotate later
    // 2 -> 90, 3 -> 180, 4 -> 45, 5 -> 180, 6 -> 30
    function applyDefaultRotationWhenEnablingRotate() {
        try {
            var sidesNow = getSelectedSideValue(radios, customInput);
            var superEff = (superEllipseCheck.value && sidesNow === 0);
            var n = getCircleAnchorCountFromRadios();
            if (!n || n < 2) n = 2;
            if (sidesNow === 0 && !superEff) {
                var ang;
                if (n === 2) ang = 90;
                else if (n === 3) ang = 180;
                else if (n === 4) ang = 45;
                else if (n === 5) ang = 180;
                else if (n === 6) ang = 30;
                else ang = 45;
                rotateInput.text = formatAngle(ang);
            } else {
                applyAutoRotationForSides(sidesNow);
                // If triangle (3 sides), default to "Down" when enabling rotate (60°)
                if (sidesNow === 3) {
                    try {
                        triangleDownRadio.value = true;
                    } catch (e) { }
                }
            }
        } catch (e) { }
    }

    function forceRotateOff() {
        rotateCheck.value = false;
        rotateInput.enabled = false;
        rotateLabel.enabled = false;
    }

    splitAtAnchorsCheck.onClick = function () {
        refreshLiveShapeAvailabilityFromUI();
        updateSuperellipseSmoothnessEnabled(getSelectedSideValue(radios, customInput));
        updatePreview();
    };

    var triangleRow = trianglePanel.add("group");
    triangleRow.orientation = "row";
    triangleRow.alignChildren = ["left", "center"];
    triangleRow.spacing = 10;

    var triangleRightRadio = triangleRow.add("radiobutton", undefined, LABELS.triangleRight[lang]);
    var triangleLeftRadio = triangleRow.add("radiobutton", undefined, LABELS.triangleLeft[lang]);
    var triangleDownRadio = triangleRow.add("radiobutton", undefined, LABELS.triangleDown[lang]);
    triangleRightRadio.value = true;

    function onTriangleDirectionChange() {
        // Ensure rotation is enabled when changing triangle direction
        rotateCheck.value = true;
        rotateInput.enabled = true;
        rotateLabel.enabled = true;
        updatePreview();
    }

    triangleRightRadio.onClick = onTriangleDirectionChange;
    triangleLeftRadio.onClick = onTriangleDirectionChange;
    triangleDownRadio.onClick = onTriangleDirectionChange;



    // Enable/disable star and pentagram options
    function validateStarAndPentagram() {
        // If Star panel is disabled (Circle), force star options off
        if (!starPanel.enabled) {
            starCheck.enabled = false;
            starCheck.value = false;
            pentagramCheck.value = false;
            pentagramCheck.enabled = false;
            return;
        }

        starCheck.enabled = true;
        if (!starCheck.value) pentagramCheck.value = false;
        pentagramCheck.enabled = starCheck.value;

        // Reuleaux is not compatible with Star shapes
        if (starCheck.value) {
            reuleauxCheck.value = false;
            reuleauxCheck.enabled = false;
        }
        // Also refresh Reuleaux amount UI when Star is ON and Reuleaux is disabled
        try { updateReuleauxAmountEnabled(); } catch (e) { }

        if (pentagramCheck.value) {
            for (var i = 0; i < 6; i++) radios[i].value = false;
            radios[3].value = true;
            customInput.enabled = false;
            applyAutoRotationForSides(5);
            forceRotateOff();
        }

        // Fallback: re-enable Reuleaux when Star is OFF (respecting odd-side rule)
        if (!starCheck.value) {
            try {
                updateReuleauxAvailability(getSelectedSideValue(radios, customInput));
            } catch (e) { }
        }
        updateInnerRadiusEnabled();
    }

    // Build current parameters from UI (used for preview and final)
    function getCurrentParams() {
        validateStarAndPentagram();

        var sides = getSelectedSideValue(radios, customInput);
        trianglePanel.enabled = (sides === 3);
        updateCirclePanelEnabled(sides);
        updateStarPanelEnabled(sides);
        updateReuleauxAvailability(sides);

        var size = parseFloat(sizeInput.text) * unitFactor;
        var ratio = parseFloat(innerRatioInput.text);
        var isStar = starCheck.value;
        var isPenta = pentagramCheck.value;
        var rotate = rotateCheck.value;
        var angle = parseFloat(rotateInput.text);
        var splitAtAnchors = splitAtAnchorsCheck.value;
        var superEllipse = superEllipseCheck.value && (sides === 0);


        var reuleaux = reuleauxCheck.value;
        var reuleauxAmount = clampReuleauxAmount(reuleauxAmountInput.text) / 100;

        var superExp = clampSuperExp(superExpInput.text);

        // Circle anchor count (effective only for Circle and not when Superellipse is ON)
        var nAnchors = getCircleAnchorCountFromRadios();
        if (!nAnchors || nAnchors < 2) nAnchors = 2;
        // Only use custom anchors when Circle is selected and Superellipse is not effective
        var circleAnchorCount = (sides === 0 && !superEllipse) ? nAnchors : 4;

        // Live shape should be disabled when using a non-default anchor count
        var isCustomCircleAnchors = (sides === 0 && !superEllipse && nAnchors !== 4);

        updateLiveShapeAvailability(splitAtAnchors, superEllipse, isCustomCircleAnchors, reuleaux);
        // Keep UI consistent
        refreshLiveShapeAvailabilityFromUI();

        // Superellipse (Circle) forces Rotate OFF
        if (superEllipse) {
            forceRotateOff();
        }

        // Auto-rotation only when Rotate is OFF (manual mode keeps user-entered angle)
        if (!rotate) {
            if (sides === 0) {
                angle = 45;
                rotateInput.text = "45";
            } else if (sides >= 3) {
                angle = 360 / (sides * 2);
                rotateInput.text = formatAngle(angle);
            }
        }

        // Triangle-specific rotation:
        // Right = -90°, Left = 90°, Down = 60°
        if (sides === 3 && trianglePanel.enabled) {
            if (triangleRightRadio.value) {
                angle = -90; // swapped
            } else if (triangleLeftRadio.value) {
                angle = 90; // swapped
            } else if (triangleDownRadio.value) {
                angle = 60;
            }
            rotateInput.text = formatAngle(angle);
        }

        if (isStar && isPenta && sides === 5) {
            ratio = (3 - Math.sqrt(5)) / 2 * 100;
            innerRatioInput.text = ratio.toFixed(2);
        }

        return {
            size: size,
            sides: sides,
            isStar: isStar,
            ratio: ratio,
            rotate: rotate,
            angle: angle,
            splitAtAnchors: splitAtAnchors,
            superEllipse: superEllipse,
            superExp: superExp,
            circleAnchorCount: circleAnchorCount,
            reuleaux: reuleaux,
            reuleauxAmount: reuleauxAmount
        };
    }

    // Update preview shape based on inputs (Undo-safe)
    function updatePreview() {
        // Always rollback the previous preview before applying a new one
        previewMgr.rollback();
        previewShape = null;

        var p = getCurrentParams();
        if (!isNaN(p.size) && !isNaN(p.ratio)) {
            previewMgr.addStep(function () {
                previewShape = createShape(app.activeDocument, p.size, p.sides, p.isStar, p.ratio, p.rotate, p.angle, p.splitAtAnchors, p.superEllipse, p.superExp, p.circleAnchorCount, p.reuleaux, p.reuleauxAmount);
            });
        }
    }

    superExpInput.onChanging = function () {
        syncSuperExpUI(superExpInput.text);
        updatePreview();
    };

    superExpSlider.onChanging = function () {
        syncSuperExpUI(superExpSlider.value);
        updatePreview();
    };

    // Event bindings
    starCheck.onClick = function () {
        validateStarAndPentagram();
        updateInnerRadiusEnabled();
        updatePreview();
    };
    pentagramCheck.onClick = function () {
        if (pentagramCheck.value) {
            forceRotateOff();
        }
        updatePreview();
    };
    superEllipseCheck.onClick = function () {
        // Superellipse is only effective for Circle (sides=0)
        var sidesNow = getSelectedSideValue(radios, customInput);
        if (superEllipseCheck.value && sidesNow === 0) {
            forceRotateOff();
        }
        try {
            var sidesNow = getSelectedSideValue(radios, customInput);
            var _superEff = (superEllipseCheck.value && sidesNow === 0);
            circleAnchorRadios.r2.enabled = !_superEff;
            circleAnchorRadios.r3.enabled = !_superEff;
            circleAnchorRadios.r4.enabled = !_superEff;
            circleAnchorRadios.r5.enabled = !_superEff;
            circleAnchorRadios.r6.enabled = !_superEff;
            // anchorCountLabel.enabled = !_superEff;
        } catch (e) { }
        refreshLiveShapeAvailabilityFromUI();
        updateSuperellipseSmoothnessEnabled(getSelectedSideValue(radios, customInput));
        updatePreview();
    };
    innerRatioInput.onChanging = function () {
        if (pentagramCheck.value) pentagramCheck.value = false;
        var v = Math.round(Number(innerRatioInput.text));
        if (isNaN(v)) v = 30;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        innerRatioInput.text = String(v);
        innerRatioSlider.value = v;
        updatePreview();
    };

    innerRatioSlider.onChanging = function () {
        var v = Math.round(innerRatioSlider.value);
        innerRatioInput.text = String(v);
        updatePreview();
    };
    sizeInput.onChanging = updatePreview;

    reuleauxCheck.onClick = function () {
        if (reuleauxCheck.value) {
            // Reset to default amount (100%) whenever Reuleaux is enabled
            syncReuleauxAmountUI(100);
        }
        updateReuleauxAmountEnabled();
        updatePreview();
    };

    reuleauxAmountInput.onChanging = function () {
        syncReuleauxAmountUI(reuleauxAmountInput.text);
        updatePreview();
    };

    reuleauxAmountSlider.onChanging = function () {
        syncReuleauxAmountUI(reuleauxAmountSlider.value);
        updatePreview();
    };

    rotateInput.onChanging = updatePreview;
    rotateCheck.onClick = function () {
        rotateInput.enabled = rotateCheck.value;
        rotateLabel.enabled = rotateCheck.value;
        if (rotateCheck.value) {
            applyDefaultRotationWhenEnablingRotate();
        }
        updatePreview();
    };
    customInput.onChanging = function () {
        var v = Math.round(Number(customInput.text));
        if (isNaN(v)) v = 12;
        if (v < 3) v = 3;
        if (v > 36) v = 36;
        customInput.text = String(v);
        customSlider.value = v;
        try { updateReuleauxAvailability(getSelectedSideValue(radios, customInput)); } catch (e) { }
        updatePreview();
    };

    customSlider.onChanging = function () {
        var v = Math.round(customSlider.value);
        customInput.text = String(v);
        try { updateReuleauxAvailability(getSelectedSideValue(radios, customInput)); } catch (e) { }
        updatePreview();
    };
    function onCircleAnchorRadioClick() {
        refreshLiveShapeAvailabilityFromUI();
        updatePreview();
    }
    circleAnchorRadios.r2.onClick = onCircleAnchorRadioClick;
    circleAnchorRadios.r3.onClick = onCircleAnchorRadioClick;
    circleAnchorRadios.r4.onClick = onCircleAnchorRadioClick;
    circleAnchorRadios.r5.onClick = onCircleAnchorRadioClick;
    circleAnchorRadios.r6.onClick = onCircleAnchorRadioClick;

    for (var i = 0; i <= 6; i++) {
        (function (i) {
            radios[i].onClick = function () {
                if (pentagramCheck.value) pentagramCheck.value = false;
                if (i === 6) {
                    for (var j = 0; j < 6; j++) radios[j].value = false;
                    customInput.enabled = true;
                    customSlider.enabled = true;
                } else {
                    radios[6].value = false;
                    customInput.enabled = false;
                    customSlider.enabled = false;
                }
                // When sides change, update rotation value even if Rotate is ON
                try { applyAutoRotationForSides(getSelectedSideValue(radios, customInput)); } catch (e) { }
                try { updateReuleauxAvailability(getSelectedSideValue(radios, customInput)); } catch (e) { }
                refreshLiveShapeAvailabilityFromUI();
                updateSuperellipseSmoothnessEnabled(getSelectedSideValue(radios, customInput));
                updatePreview();
            };
        })(i);
    }

    // Restore / Save UI state (session-only)
    function applyStateToUI(st) {
        if (!st) return;

        // Sides selection
        if (typeof st.selectedSideIndex === "number" && st.selectedSideIndex >= 0 && st.selectedSideIndex < radios.length) {
            for (var i = 0; i < radios.length; i++) radios[i].value = false;
            radios[st.selectedSideIndex].value = true;
            if (st.selectedSideIndex === 6) {
                customInput.enabled = true;
                if (typeof st.customSidesText === "string") customInput.text = st.customSidesText;
            } else {
                customInput.enabled = false;
            }
        }

        // Width
        if (typeof st.sizeText === "string") sizeInput.text = st.sizeText;

        // Rotation
        if (typeof st.rotateCheck === "boolean") rotateCheck.value = st.rotateCheck;
        if (typeof st.rotateText === "string") rotateInput.text = st.rotateText;
        rotateInput.enabled = rotateCheck.value;
        rotateLabel.enabled = rotateCheck.value;

        // Star / Pentagram
        if (typeof st.starCheck === "boolean") starCheck.value = st.starCheck;
        if (typeof st.pentagramCheck === "boolean") pentagramCheck.value = st.pentagramCheck;
        if (typeof st.innerRatioText === "string") innerRatioInput.text = st.innerRatioText;
        // Restore slider value to match input
        try {
            var v = Math.round(Number(innerRatioInput.text));
            if (!isNaN(v)) innerRatioSlider.value = v;
        } catch (e) { }
        // Enable/disable inner radius UI after restoring star/pentagram
        try { updateInnerRadiusEnabled(); } catch (e) { }
        if (typeof st.superEllipseCheck === "boolean") superEllipseCheck.value = st.superEllipseCheck;
        if (typeof st.superExpText === "string") {
            superExpInput.text = st.superExpText;
            superExpSlider.value = clampSuperExp(st.superExpText);
        }
        if (typeof st.circleAnchorsValue === "number") {
            circleAnchorRadios.r2.value = (st.circleAnchorsValue === 2);
            circleAnchorRadios.r3.value = (st.circleAnchorsValue === 3);
            circleAnchorRadios.r4.value = (st.circleAnchorsValue === 4);
            circleAnchorRadios.r5.value = (st.circleAnchorsValue === 5);
            circleAnchorRadios.r6.value = (st.circleAnchorsValue === 6);
            if (!circleAnchorRadios.r2.value && !circleAnchorRadios.r3.value && !circleAnchorRadios.r4.value && !circleAnchorRadios.r5.value && !circleAnchorRadios.r6.value) {
                circleAnchorRadios.r4.value = true;
            }
        }
        // Triangle direction
        if (st.triangleDir === "right") {
            triangleRightRadio.value = true;
        } else if (st.triangleDir === "left") {
            triangleLeftRadio.value = true;
        } else if (st.triangleDir === "down") {
            triangleDownRadio.value = true;
        }

        // Options
        // Always start with Split-at-Anchors OFF on dialog open (do not restore this from session state)
        splitAtAnchorsCheck.value = false;
        if (typeof st.reuleauxCheck === "boolean") reuleauxCheck.value = st.reuleauxCheck;
        if (typeof st.liveShapeCheck === "boolean") liveShapeCheck.value = st.liveShapeCheck;

        // Enforce split/live-shape dependency (+ Superellipse / custom anchors)
        try {
            refreshLiveShapeAvailabilityFromUI();
        } catch (e) {
            // fallback to previous behavior
            if (splitAtAnchorsCheck.value) {
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            } else {
                liveShapeCheck.enabled = true;
            }
        }

        // Enforce pentagram dependency
        if (!starCheck.value) pentagramCheck.value = false;

        // If restored state has Pentagram or Superellipse effective, force Rotate OFF
        try {
            var sidesNow = getSelectedSideValue(radios, customInput);
            if (pentagramCheck.value || (superEllipseCheck.value && sidesNow === 0)) {
                forceRotateOff();
            }
        } catch (e) { }

        // Enforce dimming of anchor radios on restore
        try {
            var _sidesNowA = getSelectedSideValue(radios, customInput);
            var _superEffA = (superEllipseCheck.value && _sidesNowA === 0);
            circleAnchorRadios.r2.enabled = !_superEffA;
            circleAnchorRadios.r3.enabled = !_superEffA;
            circleAnchorRadios.r4.enabled = !_superEffA;
            circleAnchorRadios.r5.enabled = !_superEffA;
            circleAnchorRadios.r6.enabled = !_superEffA;
            // anchorCountLabel.enabled = !_superEffA;
        } catch (e) { }

        // Update panel enabled states based on current selection
        try { updateCirclePanelEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateStarPanelEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateReuleauxAvailability(getSelectedSideValue(radios, customInput)); } catch (e) { }
    }

    function saveStateFromUI(st) {
        if (!st) return;

        // Selected side index
        var idx = 0;
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) { idx = i; break; }
        }
        st.selectedSideIndex = idx;
        st.customSidesText = customInput.text;

        // Text values
        st.sizeText = sizeInput.text;
        st.rotateCheck = rotateCheck.value;
        st.rotateText = rotateInput.text;
        st.starCheck = starCheck.value;
        st.pentagramCheck = pentagramCheck.value;
        st.innerRatioText = innerRatioInput.text;
        st.superEllipseCheck = superEllipseCheck.value;
        st.superExpText = superExpInput.text;
        st.circleAnchorsValue = getCircleAnchorCountFromRadios();

        // Triangle direction
        st.triangleDir = triangleRightRadio.value ? "right" : (triangleLeftRadio.value ? "left" : "down");

        // Options
        // This option is always reset to OFF on next open, but we still save the current value for the current session run
        st.splitAtAnchorsCheck = splitAtAnchorsCheck.value;
        st.liveShapeCheck = liveShapeCheck.value;

        st.reuleauxCheck = reuleauxCheck.value;
    }

    // Apply restored state now (before first preview)
    try { applyStateToUI(getSessionState()); } catch (e) { }
    try { updateReuleauxAmountEnabled(); } catch (e) { }
    try { updateInnerRadiusEnabled(); } catch (e) { }
    try { updateSuperellipseSmoothnessEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }

    // Save state whenever the dialog closes (OK or Cancel)
    dlg.onClose = function () {
        // Save dialog window position for this Illustrator session
        try { saveDialogLocation(dlg); } catch (_) { }
        // Save UI state
        try { saveStateFromUI(getSessionState()); } catch (e) { }

        // If cancelled, rollback preview so Undo history stays clean
        if (!confirmed) {
            try { previewMgr.rollback(); } catch (e) { }
            previewShape = null;
            // Restore view zoom on cancel
            try { if (__initView && __initZoom != null) __initView.zoom = __initZoom; } catch (_) { }
        }
    };

    dlg.onShow = function () {
        // Apply opacity first (some environments may ignore this property)
        setDialogOpacity(dlg, dialogOpacity);
        // Always start with Split-at-Anchors OFF each time the dialog opens
        try { splitAtAnchorsCheck.value = false; } catch (_) { }
        try { refreshLiveShapeAvailabilityFromUI(); } catch (_) { }

        try { updateCirclePanelEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateSuperellipseSmoothnessEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateStarPanelEnabled(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateReuleauxAvailability(getSelectedSideValue(radios, customInput)); } catch (e) { }
        try { updateReuleauxAmountEnabled(); } catch (e) { }
        try { updateInnerRadiusEnabled(); } catch (e) { }
        updatePreview();

        // Restore dialog position within this Illustrator session
        var savedLoc = null;
        try { savedLoc = getSavedDialogLocation(); } catch (_) { savedLoc = null; }
        if (savedLoc) {
            try { dlg.location = savedLoc; } catch (_) { }
        } else {
            dlg.center();
            shiftDialogPosition(dlg, offsetX, offsetY);
        }
    };

    /* ズーム / Zoom */
    var __initView = null;
    var __initZoom = null;

    var gZoom = dlg.add("group");
    gZoom.orientation = "row";
    gZoom.alignChildren = ["center", "center"];
    gZoom.alignment = "center";
    // 画面ズームの下に余白を追加
    try { gZoom.margins = [0, 7, 0, 5]; } catch (_) { }

    var stZoom = gZoom.add("statictext", undefined, LABELS.viewZoom[lang]);

    try { __initZoom = (doc && doc.activeView) ? Number(doc.activeView.zoom) : 1; } catch (_) { __initZoom = 1; }
    var sldZoom = gZoom.add("slider", undefined, (__initZoom != null ? __initZoom : 1), 0.1, 16);
    sldZoom.preferredSize.width = 240;

    function applyZoom(z) {
        try {
            if (!__initView) __initView = doc.activeView;
            if (!__initView) return;
            __initView.zoom = z;
        } catch (_) { }
    }

    sldZoom.onChanging = function () {
        applyZoom(Number(sldZoom.value));
        try { app.redraw(); } catch (_) { }
    };

    var btnArea = dlg.add("group");
    btnArea.orientation = "row";
    btnArea.alignChildren = ["center", "center"];
    btnArea.alignment = "center";
    btnArea.margins = [0, 10, 0, 0];
    btnArea.spacing = 10;

    var btnCancel = btnArea.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var btnOK = btnArea.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var confirmed = false;
    btnCancel.onClick = function () {
        // Cancel: rollback preview changes and close
        try { previewMgr.rollback(); } catch (e) { }
        previewShape = null;
        // Restore view zoom on cancel
        try { if (__initView && __initZoom != null) __initView.zoom = __initZoom; } catch (_) { }
        dlg.close();
    };
    btnOK.onClick = function () {
        // Capture current params; final shape will be applied after closing
        confirmed = true;

        dlg.close();
    };

    dlg.show();

    // If OK was pressed, finalize as a single undoable action
    if (confirmed) {
        var pFinal;
        try { pFinal = getCurrentParams(); } catch (e) { pFinal = null; }
        previewMgr.confirm(function () {
            if (!pFinal) return;
            if (isNaN(pFinal.size) || isNaN(pFinal.ratio)) return;
            previewShape = createShape(app.activeDocument, pFinal.size, pFinal.sides, pFinal.isStar, pFinal.ratio, pFinal.rotate, pFinal.angle, pFinal.splitAtAnchors, pFinal.superEllipse, pFinal.superExp, pFinal.circleAnchorCount, pFinal.reuleaux, pFinal.reuleauxAmount);
        });
    }

    if (!confirmed || !previewShape) {
        // Ensure no preview changes remain on cancel
        try { previewMgr.rollback(); } catch (e) { }
        previewShape = null;
        return null;
    }

    applyLiveShape = liveShapeCheck.value;

    return true;
}

main();
