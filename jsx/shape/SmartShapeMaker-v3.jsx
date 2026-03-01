#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SmartShapeMaker.jsx

Illustrator script to create custom shapes from a single dialog (Circle / Polygon / Star).
Real-time preview, adjustable sides, width, rotation, and options are supported.
Japanese / English UI.

Main Features:
- Specify number of sides (0 = Circle, 3/4/5/6/8, or custom)
- Star option + Pentagram option (automatic inner ratio for pentagram)
- Triangle direction options (Left / Right / Down) when sides = 3
- Width (size) input with unit display (moved into the Width panel)
- Rotation: auto angle when OFF, manual angle when ON (arrow-key editing supported)
- Options:
  - Live Shape conversion (Convert to Shape) on finalize
  - Split at Anchor Points:
    split the created path into open segments (stroke: black 0.3pt, RGB/CMYK supported),
    and automatically disables Live Shape (dimmed)

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
1. Set sides, width, star options, rotation, and options in the dialog
2. Preview updates in real-time
3. Click OK to finalize the preview object at the artboard center

Original Idea: Seiji Miyazawa (Sankai Lab)
Version: v1.2.0 (20260103)
*/

// Language detection
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "図形の作成 v1.2", en: "Create Shape v1.2" },
    shapeType: { ja: "辺の数", en: "Sides" },
    circle: { ja: "円", en: "Circle" },
    custom: { ja: "それ以外", en: "Other" },
    starPanel: { ja: "スター", en: "Star" },
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
    splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at Anchor Points" },
    liveShape: { ja: "ライブシェイプ化", en: "Live Shape" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};

var previewShape = null;
var applyLiveShape = true;

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var unit = getRulerUnitInfo();
    var result = showInputDialog(unit.label, unit.factor);
    if (!result) return;
    finalizeShape(doc);
    if (applyLiveShape) app.executeMenuCommand('Convert to Shape');
}

// Adjust value with arrow keys (Up/Down)
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
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
            try { editText.onChanging(); } catch (e) {}
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
function createShape(doc, sizePt, sides, isStar, innerRatio, rotateEnabled, rotateAngle, splitAtAnchors) {
    var layer = doc.activeLayer;
    layer.locked = false;
    layer.visible = true;
    var center = doc.activeView.centerPoint;
    var radius = sizePt / 2;
    var innerRadius = radius * (innerRatio / 100);
    var shape;

    if (sides === 0) {
        shape = layer.pathItems.ellipse(center[1] + radius, center[0] - radius, sizePt, sizePt);
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
        seg.pathPoints[0].leftDirection  = p0.anchor;
        seg.pathPoints[0].rightDirection = p0.rightDirection;
        seg.pathPoints[0].pointType = p0.pointType;

        seg.pathPoints[1].leftDirection  = p1.leftDirection;
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
        } catch (e) {}
        try { seg.strokeWidth = 0.3; } catch (e) {}
    }

    // Remove original path
    try { pathItem.remove(); } catch (e) {}

    return group;
}

// Show input dialog and handle UI and events
function showInputDialog(unitLabel, unitFactor) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    // Keyboard shortcuts
    dlg.addEventListener("keydown", function (e) {
        if (!e || !e.keyName) return;

        switch (e.keyName.toUpperCase()) {

            case "E":
                // 0 sides (Circle)
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[0].value = true;
                customInput.enabled = false;
                updatePreview();
                e.preventDefault();
                break;

            // Triangle directions and split toggle (inserted here)
            case "L":
                // Triangle Left
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
                triangleLeftRadio.value = true;
                onTriangleDirectionChange();
                e.preventDefault();
                break;

            case "R":
                // Triangle Right
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
                triangleRightRadio.value = true;
                onTriangleDirectionChange();
                e.preventDefault();
                break;

            case "B":
                // Triangle Down (Bottom)
                for (var i = 0; i < radios.length; i++) radios[i].value = false;
                radios[1].value = true; // sides = 3
                customInput.enabled = false;
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
    radios[2].value = true;

    // Rotation panel placed under the sides panel (left column)
    var rotatePanel = leftCol.add("group");
    rotatePanel.orientation = "row";
    rotatePanel.alignChildren = "center";
    rotatePanel.alignment = "center";
    rotatePanel.margins = [0, 6, 0, 0];

    var rotateCheck = rotatePanel.add("checkbox", undefined, LABELS.rotation[lang]);
    var rotateInput = rotatePanel.add("edittext", undefined, "90");
    rotateInput.characters = 4;
    changeValueByArrowKey(rotateInput);
    var rotateLabel = rotatePanel.add("statictext", undefined, "°");
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

    var starCheck = starPanel.add("checkbox", undefined, LABELS.star[lang]);

    var innerGroup = starPanel.add("group");
    innerGroup.add("statictext", undefined, LABELS.innerRadius[lang]);
    var innerRatioInput = innerGroup.add("edittext", undefined, "30");
    innerRatioInput.characters = 4;
    changeValueByArrowKey(innerRatioInput);
    innerGroup.add("statictext", undefined, LABELS.percent[lang]);

    var pentagramCheck = starPanel.add("checkbox", undefined, LABELS.pentagram[lang]);
    pentagramCheck.value = false;

    // Triangle panel placed under the Star panel
    var trianglePanel = right.add("panel", undefined, LABELS.trianglePanel[lang]);
    trianglePanel.orientation = "column";
    trianglePanel.alignChildren = "left";
    trianglePanel.margins = [15, 20, 15, 10];

    // Options panel placed under the Triangle panel
    var optionPanel = right.add("panel", undefined, LABELS.optionPanel[lang]);
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.margins = [15, 20, 15, 10];

    // Live shape option moved here
    var liveShapeCheck = optionPanel.add("checkbox", undefined, LABELS.liveShape[lang]);
    liveShapeCheck.value = true;

    // Split at anchor points
    var splitAtAnchorsCheck = optionPanel.add("checkbox", undefined, LABELS.splitAtAnchors[lang]);
    splitAtAnchorsCheck.value = false;
    splitAtAnchorsCheck.onClick = function () {
        if (splitAtAnchorsCheck.value) {
            liveShapeCheck.value = false;
            liveShapeCheck.enabled = false; // dim
        } else {
            liveShapeCheck.enabled = true;
        }
        updatePreview();
    };

    var triangleRightRadio = trianglePanel.add("radiobutton", undefined, LABELS.triangleRight[lang]);
    var triangleLeftRadio  = trianglePanel.add("radiobutton", undefined, LABELS.triangleLeft[lang]);
    var triangleDownRadio  = trianglePanel.add("radiobutton", undefined, LABELS.triangleDown[lang]);
    triangleRightRadio.value = true;

    function onTriangleDirectionChange() {
        // Ensure rotation is enabled when changing triangle direction
        rotateCheck.value = true;
        rotateInput.enabled = true;
        rotateLabel.enabled = true;
        updatePreview();
    }

    triangleRightRadio.onClick = onTriangleDirectionChange;
    triangleLeftRadio.onClick  = onTriangleDirectionChange;
    triangleDownRadio.onClick  = onTriangleDirectionChange;



    // Enable/disable star and pentagram options
    function validateStarAndPentagram() {
        starCheck.enabled = true;
        if (!starCheck.value) pentagramCheck.value = false;
        pentagramCheck.enabled = starCheck.value;
        if (pentagramCheck.value) {
            for (var i = 0; i < 6; i++) radios[i].value = false;
            radios[3].value = true;
            customInput.enabled = false;
        }
    }

    // Update preview shape based on inputs
    function updatePreview() {
        validateStarAndPentagram();

        var sides = getSelectedSideValue(radios, customInput);
        trianglePanel.enabled = (sides === 3);

        var size = parseFloat(sizeInput.text) * unitFactor;
        var ratio = parseFloat(innerRatioInput.text);
        var isStar = starCheck.value;
        var isPenta = pentagramCheck.value;
        var rotate = rotateCheck.value;
        var angle = parseFloat(rotateInput.text);
        var splitAtAnchors = splitAtAnchorsCheck.value;

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

        if (!isNaN(size) && !isNaN(ratio)) {
            if (previewShape) try { previewShape.remove(); } catch (e) {}
            previewShape = createShape(app.activeDocument, size, sides, isStar, ratio, rotate, angle, splitAtAnchors);
            app.redraw();
        }
    }

    // Event bindings
    starCheck.onClick = updatePreview;
    pentagramCheck.onClick = updatePreview;
    innerRatioInput.onChanging = function () {
        if (pentagramCheck.value) pentagramCheck.value = false;
        updatePreview();
    };
    sizeInput.onChanging = updatePreview;
    rotateInput.onChanging = updatePreview;
    rotateCheck.onClick = function () {
        rotateInput.enabled = rotateCheck.value;
        rotateLabel.enabled = rotateCheck.value;
        updatePreview();
    };
    customInput.onChanging = updatePreview;

    for (var i = 0; i <= 6; i++) {
        (function (i) {
            radios[i].onClick = function () {
                if (pentagramCheck.value) pentagramCheck.value = false;
                if (i === 6) {
                    for (var j = 0; j < 6; j++) radios[j].value = false;
                    customInput.enabled = true;
                } else {
                    radios[6].value = false;
                    customInput.enabled = false;
                }
                updatePreview();
            };
        })(i);
    }

    dlg.addEventListener("show", function () {
        updatePreview();
        dlg.center();
        dlg.location = [dlg.location[0] + 300, dlg.location[1]];
    });

    var btnArea = dlg.add("group");
    btnArea.orientation = "row";
    btnArea.alignChildren = ["center", "center"];
    btnArea.alignment = "center";
    btnArea.margins = [0, 10, 0, 0];
    btnArea.spacing = 10;

    var btnCancel = btnArea.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var btnOK = btnArea.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var confirmed = false;
    btnCancel.onClick = function () { dlg.close(); };
    btnOK.onClick = function () {
        // Ensure latest UI state is reflected in the preview object before closing
        try { updatePreview(); } catch (e) {}

        confirmed = true;
        dlg.close();
    };

    dlg.show();

    if (!confirmed || !previewShape) {
        if (previewShape) try { previewShape.remove(); } catch (e) {}
        previewShape = null;
        return null;
    }

    applyLiveShape = liveShapeCheck.value;

    return {
        size: parseFloat(sizeInput.text),
        sides: getSelectedSideValue(radios, customInput),
        rotateEnabled: rotateCheck.value,
        rotateAngle: parseFloat(rotateInput.text)
    };
}

main();
