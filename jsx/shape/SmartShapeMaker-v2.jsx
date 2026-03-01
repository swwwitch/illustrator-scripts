#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SmartShapeMaker.jsx

Illustrator script to create custom shapes by combining Rectangle, Ellipse, Polygon, and Star tools into one dialog.
Features include real-time preview, adjustable sides, radii, rotation angle, and live shape option.
Supports Japanese and English UI.

Main Features:
- Specify number of sides (including circle and custom)
- Star and pentagram options
- Adjust outer radius and inner radius (percentage)
- Automatic or manual rotation angle
- Specify width (size) and option to create live shape

Usage Flow:
1. Set sides, size, radius ratio, rotation, and options in dialog
2. Preview shape in real-time
3. Click OK to draw shape at artboard center

Original Idea: Seiji Miyazawa (Sankai Lab)
Version: v1.1.0 (20250503)
*/

// Language detection
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "図形の作成 v1.1", en: "Create Shape v1.1" },
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
function createShape(doc, sizePt, sides, isStar, innerRatio, rotateEnabled, rotateAngle) {
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

    doc.selection = [shape];
    return shape;
}

// Show input dialog and handle UI and events
function showInputDialog(unitLabel, unitFactor) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var radios = [], customInput;
    var main = dlg.add("group");
    main.orientation = "row";

    var left = main.add("panel", undefined, LABELS.shapeType[lang]);
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

    // Rotation panel (moved below sides panel)
    var rotatePanel = dlg.add("group");
    rotatePanel.orientation = "row";
    rotatePanel.alignChildren = "center";
    rotatePanel.margins = [15, 10, 15, 0];

    var rotateCheck = rotatePanel.add("checkbox", undefined, LABELS.rotation[lang]);
    var rotateInput = rotatePanel.add("edittext", undefined, "90");
    rotateInput.characters = 4;
    changeValueByArrowKey(rotateInput);
    var rotateLabel = rotatePanel.add("statictext", undefined, "°");

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

    var divider = dlg.add("panel");
    divider.alignment = "fill";
    divider.minimumSize.height = 1;
    divider.margins = [0, 10, 0, 10];

    var bottomGroup = dlg.add("group");
    bottomGroup.orientation = "row";
    bottomGroup.alignChildren = ["left", "center"];
    bottomGroup.margins = [20, 0, 20, 10];

    bottomGroup.add("statictext", undefined, LABELS.width[lang]);
    bottomGroup.orientation = "row";
    bottomGroup.alignChildren = ["left", "center"];
    var sizeInput = bottomGroup.add("edittext", undefined, "100");
    sizeInput.characters = 5;
    changeValueByArrowKey(sizeInput);
    bottomGroup.add("statictext", undefined, "(" + unitLabel + ")");

    var liveShapeCheck = bottomGroup.add("checkbox", undefined, LABELS.liveShape[lang]);
    liveShapeCheck.value = true;

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

        if (sides === 0) {
            angle = 45;
            rotateInput.text = "45";
        } else if (sides >= 3) {
            angle = 360 / (sides * 2);
            rotateInput.text = formatAngle(angle);
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
            previewShape = createShape(app.activeDocument, size, sides, isStar, ratio, rotate, angle);
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
    btnArea.alignChildren = ["right", "center"];
    btnArea.alignment = "right";
    btnArea.margins = [0, 10, 0, 0];
    btnArea.spacing = 10;

    var btnCancel = btnArea.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var btnOK = btnArea.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var confirmed = false;
    btnCancel.onClick = function () { dlg.close(); };
    btnOK.onClick = function () {
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
