#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


/*
### スクリプト名：

SmartShapeMaker.jsx

### 概要

- 「長方形ツール」「楕円形ツール」「多角形ツール」「スターツール」を一つのダイアログにまとめ、自由にカスタム図形を作成するIllustrator用スクリプトです。
- リアルタイムプレビューで確認しながら、辺の数、半径、回転角度、ライブシェイプ化など多彩な設定が可能です。

### 主な機能

- 辺数指定（正多角形、円、カスタム）
- スターおよび五芒星オプション
- 外半径、内半径（比率％）の調整
- 回転角度自動計算および手動設定
- 幅（サイズ）指定、ライブシェイプ生成オプション
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログで辺の数、サイズ、半径比、回転角度、オプションを設定
2. リアルタイムプレビューで確認
3. OKを押すとアートボード中央に図形を描画

### オリジナル、謝辞

オリジナルアイデア：宮澤聖二さん（三階ラボ）

### 更新履歴

- v1.0.0 (20250503) : 初期バージョン

---

### Script Name:

SmartShapeMaker.jsx

### Overview

- An Illustrator script that combines "Rectangle Tool", "Ellipse Tool", "Polygon Tool", and "Star Tool" into a single dialog for flexible shape creation.
- Allows setting sides, radii, rotation angle, live shape option, with real-time preview.

### Main Features

- Specify number of sides (polygons, circle, custom)
- Star and pentagram options
- Adjust outer radius and inner radius (percentage)
- Automatic or manual rotation angle
- Specify width (size), option to create as live shape
- Japanese and English UI support

### Process Flow

1. Configure sides, size, radius ratio, rotation angle, and options in the dialog
2. Check in real-time preview
3. Click OK to draw the shape at the center of the artboard

### Original / Acknowledgements

Original idea: Seiji Miyazawa (Sankai Lab)

### Update History

- v1.0.0 (20250503): Initial version
*/


// -------------------------------
// 日英ラベル定義 / Define labels
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle:     { ja: "図形の作成", en: "Create Shape" },
    shapeType:       { ja: "辺の数", en: "Sides" },
    starPanel:       { ja: "スター", en: "Star" },
    outerRadius:     { ja: "第1半径：", en: "Outer Radius:" },
    innerRadius:     { ja: "第2半径：", en: "Inner Radius:" },
    percent:         { ja: "%", en: "%" },
    custom:          { ja: "それ以外", en: "Other" },
    rotation:        { ja: "回転", en: "Rotate" },
    width:           { ja: "幅：", en: "Width:" },
    unitLabel:       { ja: "（単位）", en: "(unit)" },
    liveShape:       { ja: "ライブシェイプ", en: "Live Shape" },
    ok:              { ja: "OK", en: "OK" },
    cancel:          { ja: "キャンセル", en: "Cancel" },
    pentagram:       { ja: "五芒星", en: "Pentagram" },
    star:            { ja: "スター", en: "Star" },
    circle:            { ja: "円", en: "Circle" }
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

function formatAngle(value) {
    var rounded = Math.round(value * 1000) / 1000;
    return (rounded % 1 === 0) ? String(Math.round(rounded)) : String(rounded);
}

function getSelectedSideValue(radios, input) {
    for (var i = 0; i < radios.length; i++) {
        if (radios[i].value) {
            return (i === 6) ? parseInt(input.text, 10) : [0, 3, 4, 5, 6, 8][i];
        }
    }
    return 4;
}

function finalizeShape(doc) {
    if (!previewShape) return;
    doc.selection = [previewShape];
    previewShape = null;
}

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
    radios[2].value = true;

    var rotatePanel = dlg.add("group");
    rotatePanel.orientation = "row";
    rotatePanel.alignChildren = "center";
    rotatePanel.margins = [20, 0, 20, 0];
    var rotateCheck = rotatePanel.add("checkbox", undefined, LABELS.rotation[lang]);
    var rotateInput = rotatePanel.add("edittext", undefined, "90");
    rotateInput.characters = 4;
    var rotateLabel = rotatePanel.add("statictext", undefined, "°");

    var right = main.add("panel", undefined, LABELS.starPanel[lang]);
    right.orientation = "column";
    right.alignChildren = "left";
    right.alignment = "top";
    right.margins = [20, 20, 10, 10];

    var starCheck = right.add("checkbox", undefined, LABELS.star[lang]);

    var innerGroup = right.add("group");
    innerGroup.add("statictext", undefined, LABELS.innerRadius[lang]);
    var innerRatioInput = innerGroup.add("edittext", undefined, "30");
    innerRatioInput.characters = 4;
    innerGroup.add("statictext", undefined, LABELS.percent[lang]);

    var pentagramCheck = right.add("checkbox", undefined, LABELS.pentagram[lang]);
    pentagramCheck.value = false;

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
    bottomGroup.add("statictext", undefined, "(" + unitLabel + ")");

    var liveShapeCheck = bottomGroup.add("checkbox", undefined, LABELS.liveShape[lang]);
    liveShapeCheck.value = true;

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

    function updatePreview() {
        validateStarAndPentagram();

        var sides = getSelectedSideValue(radios, customInput);
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

    // イベント設定
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
btnArea.alignChildren = ["right", "center"]; // ★ 右寄せに変更
btnArea.alignment = "right";                // ★ ダイアログ内で右寄せ
btnArea.margins = [0, 10, 0, 0];
btnArea.spacing = 10; // ボタン間のスペース


var btnCancel = btnArea.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
var btnOK     = btnArea.add("button", undefined, LABELS.ok[lang],     { name: "ok" });

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
