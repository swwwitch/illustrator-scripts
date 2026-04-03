#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * チケットメーカー3.jsx
 * 概要: 選択した長方形に対して、ミシン目・エッジ・コーナー・ホールを付加し、チケット風の形状を作成します。
 * Summary: Adds perforations, edges, corners, and holes to selected rectangles to create ticket-like shapes.
 * 更新日: 2026-03-08
 * Updated: 2026-03-08
 */

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

/* 単位ユーティリティ / Unit utilities */
var unitMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;
        case 1: return 72.0 / 25.4;
        case 2: return 1.0;
        case 3: return 12.0;
        case 4: return 72.0 / 2.54;
        case 5: return 72.0 / 25.4 * 0.25;
        case 6: return 1.0;
        case 7: return 72.0 * 12.0;
        case 8: return 72.0 / 25.4 * 1000.0;
        case 9: return 72.0 * 36.0;
        case 10: return 72.0 * 12.0;
        default: return 1.0;
    }
}

function toPt(value, factor) {
    return value * factor;
}

function fromPt(value, factor) {
    return value / factor;
}

var rulerUnitCode = app.preferences.getIntegerPreference("rulerType");
var rulerUnitLabel = getUnitLabel(rulerUnitCode, "rulerType");
var rulerPtFactor = getPtFactorFromUnitCode(rulerUnitCode);

var strokeUnitCode = app.preferences.getIntegerPreference("strokeUnits");
var strokeUnitLabel = getUnitLabel(strokeUnitCode, "strokeUnits");
var strokePtFactor = getPtFactorFromUnitCode(strokeUnitCode);

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "チケットメーカー", en: "Ticket Maker" },
    alertSelectRectangle: { ja: "長方形を選択してください。", en: "Please select a rectangle." },
    alertEnterNumbers: { ja: "数値を入力してください。", en: "Please enter numeric values." },
    previewLayerName: { ja: "プレビュー", en: "Preview" },
    panelPerforation: { ja: "ミシン目", en: "Perforation" },
    panelDash: { ja: "破線", en: "Dashed Line" },
    panelEdge: { ja: "エッジ", en: "Edge" },
    panelCorner: { ja: "コーナー", en: "Corner" },
    chkMiddle: { ja: "中央", en: "Center" },
    lblLineWidth: { ja: "線幅:", en: "Weight:" },
    lblGap: { ja: "ドット間隔:", en: "Dot Gap:" },
    lblLength: { ja: "端余白:", en: "Edge Offset:" },
    unitRuler: { ja: rulerUnitLabel, en: rulerUnitLabel },
    unitStroke: { ja: strokeUnitLabel, en: strokeUnitLabel },
    rdNone: { ja: "なし", en: "None" },
    rdCircle: { ja: "円", en: "Circle" },
    rdTriangle: { ja: "三角", en: "Triangle" },
    rdRound: { ja: "角丸", en: "Rounded" },
    rdInverse: { ja: "逆角丸", en: "Inverse Round" },
    rdChamfer: { ja: "面取り", en: "Chamfer" },
    panelHole: { ja: "スリット／ホール", en: "Slit / Hole" },
    chkLeftRight: { ja: "左右にも追加", en: "Add Left/Right" },
    lblSize: { ja: "サイズ:", en: "Size:" },
    chkShapeOnly: { ja: "エッジのみ", en: "Edge Only" },
    chkLeft: { ja: "左", en: "Left" },
    chkRight: { ja: "右", en: "Right" },
    chkPreview: { ja: "プレビュー", en: "Preview" },
    chkExpandAppearance: { ja: "アピアランスを分割", en: "Expand Appearance" },
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOK: { ja: "OK", en: "OK" }
};

function L(key) {
    var item = LABELS[key];
    return item ? item[lang] : key;
}
function ensureStrokeDotActionLoaded() {
    if (ensureStrokeDotActionLoaded._loaded) return;

    var str = '/version 3 /name [ 9 5374726f6b65446f74 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 3 646f74 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_setStroke) /localizedName [ 12 e7b79ae38292e8a8ade5ae9a ] /isOpen 1 /isOn 1 /hasDialog 0 /parameterCount 11 /parameter-1 { /key 2003072104 /showInPalette 4294967295 /type (unit real) /value 6.0000944882 /unit 592476268 } /parameter-2 { /key 1667330094 /showInPalette 4294967295 /type (enumerated) /name [ 12 e4b8b8e59e8be7b79ae7abaf ] /value 1 } /parameter-3 { /key 1785686382 /showInPalette 4294967295 /type (enumerated) /name [ 15 e38399e38399e383abe7b590e59088 ] /value 2 } /parameter-4 { /key 1684825454 /showInPalette 4294967295 /type (integer) /value 2 } /parameter-5 { /key 1685284913 /showInPalette 4294967295 /type (unit real) /value 0.0 /unit 592476268 } /parameter-6 { /key 1685284914 /showInPalette 4294967295 /type (unit real) /value 19.8425197601 /unit 592476268 } /parameter-7 { /key 1684104298 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-8 { /key 1634231345 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-9 { /key 1634231346 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-10 { /key 1634230636 /showInPalette 4294967295 /type (enumerated) /name [ 24 e38391e382b9e381aee7b582e782b9e381abe9858de7bdae ] /value 0 } /parameter-11 { /key 1634494318 /showInPalette 4294967295 /type (enumerated) /name [ 6 e4b8ade5a4ae ] /value 0 } } }';

    var f = new File('~/StrokeDot.aia');
    f.open('w');
    f.write(str);
    f.close();
    app.loadAction(f);
    f.remove();

    ensureStrokeDotActionLoaded._loaded = true;
}

function unloadStrokeDotAction() {
    if (!ensureStrokeDotActionLoaded._loaded) return;
    try {
        app.unloadAction('StrokeDot', '');
    } catch (e) {
    }
    ensureStrokeDotActionLoaded._loaded = false;
}

function act_StrokeDot() {
    ensureStrokeDotActionLoaded();
    app.doScript('dot', 'StrokeDot', false);
}

var doc = app.activeDocument;
var sel = doc.selection;

if (sel.length === 0) {
    alert(L('alertSelectRectangle'));
} else {
    // app.executeMenuCommand('edge');
    try {
        var originalLayer = sel[0].layer;

        // プレビュー専用レイヤーを最前面に作成
        var previewLayer = doc.layers.add();
        previewLayer.name = L('previewLayerName');

        /* エフェクトを適用する関数 / Apply effect */
        /* isPreview=true: プレビューレイヤーに複製して適用、元を非表示 / Duplicate to preview layer and hide originals */
        /* isPreview=false: 元のオブジェクトに直接適用 / Apply directly to original objects */
        function applyEffect(isPreview) {
            var targets = [];

            for (var i = 0; i < sel.length; i++) {
                if (isPreview) {
                    doc.activeLayer = previewLayer;
                    var target = sel[i].duplicate(previewLayer, ElementPlacement.PLACEATEND);
                    sel[i].hidden = true;
                    targets.push(target);
                } else {
                    targets.push(sel[i]);
                }
            }

            var strokeWValue = parseFloat(inputWidth.text);
            var gapValue = parseFloat(inputGap.text);
            var dashLenValue = parseFloat(inputDashLen.text);
            var offsetValue = parseFloat(inputOffset.text);
            if (isNaN(strokeWValue) || isNaN(gapValue) || isNaN(dashLenValue) || isNaN(offsetValue)) return;

            var strokeW = toPt(strokeWValue, strokePtFactor);
            var gap = toPt(gapValue, rulerPtFactor);
            var dashLen = toPt(dashLenValue, rulerPtFactor);
            var offsetPt = toPt(offsetValue, rulerPtFactor);

            // 角丸を適用
            if (rdRound.value) {
                var cornerSizeValue = parseFloat(inputCornerSize.text);
                if (!isNaN(cornerSizeValue) && cornerSizeValue > 0) {
                    var cornerSizePt = toPt(cornerSizeValue, rulerPtFactor);
                    var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSizePt + ' "/></LiveEffect>';
                    for (var i = 0; i < targets.length; i++) {
                        targets[i].applyEffect(xml);
                    }
                }
            }

            for (var i = 0; i < targets.length; i++) {
                var rect = targets[i];
                var x = rect.left;
                var y = rect.top;
                var w = rect.width;
                var h = rect.height;
                var lineX = x + w / 2 + offsetPt;
                var items = [rect];

                // ミシン目を作成（中間）
                if (chkMiddle.value && !(chkEdgeOnly.value && !rdEdgeNone.value)) {
                    var line = doc.pathItems.add();
                    line.setEntirePath([[lineX, y - dashLen], [lineX, y - h + dashLen]]);
                    line.filled = false;
                    line.stroked = true;

                    doc.selection = null;
                    line.selected = true;
                    act_StrokeDot();
                    line.strokeWidth = strokeW;
                    line.strokeDashes = [0, gap];
                    doc.selection = null;
                    line.selected = true;
                    app.executeMenuCommand('Live Outline Stroke');
                    items.push(doc.selection[0]);
                }

                // ミシン目を作成（左右）
                if (chkLR.value) {
                    var lrPositions = [x, x + w];
                    for (var lr = 0; lr < lrPositions.length; lr++) {
                        var lrLine = doc.pathItems.add();
                        lrLine.setEntirePath([[lrPositions[lr], y - dashLen], [lrPositions[lr], y - h + dashLen]]);
                        lrLine.filled = false;
                        lrLine.stroked = true;

                        doc.selection = null;
                        lrLine.selected = true;
                        act_StrokeDot();
                        lrLine.strokeWidth = strokeW;
                        lrLine.strokeDashes = [0, gap];
                        doc.selection = null;
                        lrLine.selected = true;
                        app.executeMenuCommand('Live Outline Stroke');
                        items.push(doc.selection[0]);
                    }
                }

                // エッジを作成
                if (!rdEdgeNone.value) {
                    var edgeSizeValue = parseFloat(inputEdgeSize.text);
                    if (!isNaN(edgeSizeValue) && edgeSizeValue > 0) {
                        var edgeSizePt = toPt(edgeSizeValue, strokePtFactor);
                        var edgeR = edgeSizePt / 2;
                        var edgeXList = [];
                        if (chkMiddle.value) edgeXList.push(lineX);

                        for (var ei = 0; ei < edgeXList.length; ei++) {
                            var ex = edgeXList[ei];
                            var edgePositions = [[ex, y], [ex, y - h]];
                            for (var ep = 0; ep < edgePositions.length; ep++) {
                                var epx = edgePositions[ep][0];
                                var epy = edgePositions[ep][1];
                                if (rdEdgeCircle.value) {
                                    var edgeShape = doc.pathItems.ellipse(epy + edgeR, epx - edgeR, edgeSizePt, edgeSizePt);
                                    edgeShape.filled = true;
                                    edgeShape.stroked = false;
                                } else {
                                    var edgeShape = doc.pathItems.rectangle(epy + edgeR, epx - edgeR, edgeSizePt, edgeSizePt);
                                    edgeShape.filled = true;
                                    edgeShape.stroked = false;
                                    edgeShape.rotate(45);
                                }
                                items.push(edgeShape);
                            }
                        }
                    }
                }

                // ホールを作成
                var holeSizeValue = parseFloat(inputHoleSize.text);
                if (!rdHoleNone.value && !isNaN(holeSizeValue) && holeSizeValue > 0) {
                    var holeSizePt = toPt(holeSizeValue, rulerPtFactor);
                    var cy = y - h / 2;
                    var r = holeSizePt / 2;

                    var holeItems = [];

                    if (rdHoleCircle.value) {
                        // 円
                        if (chkHoleLeft.value) {
                            var leftCircle = doc.pathItems.ellipse(cy + r, x - r, holeSizePt, holeSizePt);
                            leftCircle.filled = true;
                            leftCircle.stroked = false;
                            holeItems.push(leftCircle);
                        }
                        if (chkHoleRight.value) {
                            var rightCircle = doc.pathItems.ellipse(cy + r, x + w - r, holeSizePt, holeSizePt);
                            rightCircle.filled = true;
                            rightCircle.stroked = false;
                            holeItems.push(rightCircle);
                        }
                    } else {
                        // 三角（45°回転した正方形）
                        if (chkHoleLeft.value) {
                            var leftRect = doc.pathItems.rectangle(cy + r, x - r, holeSizePt, holeSizePt);
                            leftRect.filled = true;
                            leftRect.stroked = false;
                            leftRect.rotate(45);
                            holeItems.push(leftRect);
                        }
                        if (chkHoleRight.value) {
                            var rightRect = doc.pathItems.rectangle(cy + r, x + w - r, holeSizePt, holeSizePt);
                            rightRect.filled = true;
                            rightRect.stroked = false;
                            rightRect.rotate(45);
                            holeItems.push(rightRect);
                        }
                    }

                    if (holeItems.length > 0) {
                        doc.selection = null;
                        for (var hi = 0; hi < holeItems.length; hi++) {
                            holeItems[hi].selected = true;
                        }
                        if (holeItems.length > 1) {
                            app.executeMenuCommand('group');
                        }
                        items.push(doc.selection[0]);
                    }
                }

                // コーナーの逆角丸を作成
                if (rdInverse.value) {
                    var cSizeValue = parseFloat(inputCornerSize.text);
                    if (!isNaN(cSizeValue) && cSizeValue > 0) {
                        var cSizePt = toPt(cSizeValue, rulerPtFactor);

                        // 上辺（左上・右上の角）
                        var topLine = doc.pathItems.add();
                        topLine.setEntirePath([[x, y], [x + w, y]]);
                        topLine.filled = false;
                        topLine.stroked = true;

                        doc.selection = null;
                        topLine.selected = true;
                        act_StrokeDot();
                        topLine.strokeWidth = cSizePt;
                        topLine.strokeDashes = [0, 1000];
                        doc.selection = null;
                        topLine.selected = true;
                        app.executeMenuCommand('Live Outline Stroke');
                        items.push(doc.selection[0]);

                        // 下辺（左下・右下の角）
                        var bottomLine = doc.pathItems.add();
                        bottomLine.setEntirePath([[x, y - h], [x + w, y - h]]);
                        bottomLine.filled = false;
                        bottomLine.stroked = true;

                        doc.selection = null;
                        bottomLine.selected = true;
                        act_StrokeDot();
                        bottomLine.strokeWidth = cSizePt;
                        bottomLine.strokeDashes = [0, 1000];
                        doc.selection = null;
                        bottomLine.selected = true;
                        app.executeMenuCommand('Live Outline Stroke');
                        items.push(doc.selection[0]);
                    }
                }

                // コーナーの面取りを作成
                if (rdChamfer.value) {
                    var cSizeValue = parseFloat(inputCornerSize.text);
                    if (!isNaN(cSizeValue) && cSizeValue > 0) {
                        var cSizePt = toPt(cSizeValue, rulerPtFactor);
                        var corners = [
                            [x, y],
                            [x + w, y],
                            [x, y - h],
                            [x + w, y - h]
                        ];
                        for (var c = 0; c < corners.length; c++) {
                            var cx = corners[c][0];
                            var cy = corners[c][1];
                            var sq = doc.pathItems.rectangle(cy + cSizePt / 2, cx - cSizePt / 2, cSizePt, cSizePt);
                            sq.filled = true;
                            sq.stroked = false;
                            sq.rotate(45);
                            items.push(sq);
                        }
                    }
                }

                // グループ化してパスファインダー適用
                doc.selection = null;
                for (var k = 0; k < items.length; k++) {
                    items[k].selected = true;
                }
                app.executeMenuCommand('group');
                app.executeMenuCommand('Live Pathfinder Subtract');
            }

            app.redraw();
        }

        /* プレビューレイヤーをクリアし元オブジェクトを復元 / Clear preview layer and restore original objects */
        function removePreviewLayer() {
            while (previewLayer.pageItems.length > 0) {
                previewLayer.pageItems[0].remove();
            }
            for (var i = 0; i < sel.length; i++) {
                sel[i].hidden = false;
            }
            app.redraw();
        }

        function changeValueByArrowKey(editText, allowNegative, targetInput, textFrames) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;
                if (event.keyName != "Up" && event.keyName != "Down") return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    if (event.keyName == "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value = Math.floor((value - 1) / delta) * delta;
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
                        event.preventDefault();
                    }
                }

                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                if (!allowNegative && value < 0) value = 0;

                editText.text = value;

                if (targetInput === inputOffset && typeof sliderOffset !== 'undefined' && sliderOffset) {
                    sliderOffset.value = Math.max(-halfW_mm, Math.min(halfW_mm, value));
                }

                if (typeof updatePreview === "function") {
                    updatePreview();
                }
            });
        }

        /* ダイアログボックス / Dialog */
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];

        var labelW = 70;

        var grpCornerHole = dlg.add('group');
        grpCornerHole.orientation = 'row';
        grpCornerHole.alignChildren = ['fill', 'top'];

        var pnlCorner = grpCornerHole.add('panel', undefined, L('panelCorner'));
        pnlCorner.orientation = 'column';
        pnlCorner.alignChildren = ['left', 'top'];
        pnlCorner.margins = [15, 20, 15, 10];

        var grpCornerRadio = pnlCorner.add('group');
        grpCornerRadio.orientation = 'column';
        grpCornerRadio.alignChildren = ['left', 'top'];
        grpCornerRadio.alignment = ['left', 'top'];
        var rdCornerNone = grpCornerRadio.add('radiobutton', undefined, L('rdNone'));
        var rdRound = grpCornerRadio.add('radiobutton', undefined, L('rdRound'));
        var rdInverse = grpCornerRadio.add('radiobutton', undefined, L('rdInverse'));
        var rdChamfer = grpCornerRadio.add('radiobutton', undefined, L('rdChamfer'));
        rdCornerNone.value = true;

        var grpCornerSize = pnlCorner.add('group');
        grpCornerSize.add('statictext', undefined, L('lblSize'));
        var inputCornerSize = grpCornerSize.add('edittext', undefined, '5');
        inputCornerSize.characters = 3;
        changeValueByArrowKey(inputCornerSize, false, inputCornerSize);
        grpCornerSize.add('statictext', undefined, L('unitRuler'));

        var pnlHole = grpCornerHole.add('panel', undefined, L('panelHole'));
        pnlHole.orientation = 'column';
        pnlHole.alignChildren = ['left', 'top'];
        pnlHole.margins = [15, 20, 15, 10];

        var grpHoleRadio = pnlHole.add('group');
        grpHoleRadio.orientation = 'column';
        grpHoleRadio.alignChildren = ['left', 'top'];
        grpHoleRadio.alignment = ['left', 'top'];
        var rdHoleNone = grpHoleRadio.add('radiobutton', undefined, L('rdNone'));
        var rdHoleCircle = grpHoleRadio.add('radiobutton', undefined, L('rdCircle'));
        var rdHoleTriangle = grpHoleRadio.add('radiobutton', undefined, L('rdTriangle'));
        rdHoleNone.value = true;

        var grpHoleSide = pnlHole.add('group');
        grpHoleSide.orientation = 'row';
        var chkHoleLeft = grpHoleSide.add('checkbox', undefined, L('chkLeft'));
        var chkHoleRight = grpHoleSide.add('checkbox', undefined, L('chkRight'));
        chkHoleLeft.value = true;
        chkHoleRight.value = true;

        var grpHoleSize = pnlHole.add('group');
        grpHoleSize.add('statictext', undefined, L('lblSize'));
        var inputHoleSize = grpHoleSize.add('edittext', undefined, '5');
        inputHoleSize.characters = 3;
        changeValueByArrowKey(inputHoleSize, false, inputHoleSize);
        grpHoleSize.add('statictext', undefined, L('unitRuler'));

        var pnl = dlg.add('panel', undefined, L('panelPerforation'));
        pnl.orientation = 'column';
        pnl.alignChildren = ['fill', 'top'];
        pnl.margins = [15, 20, 15, 10];

        var grpOffsetWrap = pnl.add('group');
        grpOffsetWrap.orientation = 'column';
        grpOffsetWrap.alignChildren = ['center', 'top'];

        var grpMishimeTop = grpOffsetWrap.add('group');
        grpMishimeTop.orientation = 'row';
        var chkLR = grpMishimeTop.add('checkbox', undefined, L('chkLeftRight'));
        var chkMiddle = grpMishimeTop.add('checkbox', undefined, L('chkMiddle'));
        chkMiddle.value = true;
        var labelW2 = 78;
        var inputOffset = grpMishimeTop.add('edittext', undefined, '0');
        inputOffset.characters = 3;
        changeValueByArrowKey(inputOffset, true, inputOffset);
        var lblOffsetMM = grpMishimeTop.add('statictext', undefined, L('unitRuler'));

        var halfW_mm = Math.round(fromPt(sel[0].width / 2, rulerPtFactor) * 10) / 10;
        var sliderOffset = grpOffsetWrap.add('slider', undefined, 0, -halfW_mm, halfW_mm);
        sliderOffset.alignment = ['center', 'top'];
        sliderOffset.preferredSize.width = 180;

        sliderOffset.onChanging = function () {
            inputOffset.text = String(Math.round(sliderOffset.value * 10) / 10);
            updatePreview();
        };
        inputOffset.onChanging = function () {
            var v = parseFloat(inputOffset.text);
            if (!isNaN(v)) {
                sliderOffset.value = Math.max(-halfW_mm, Math.min(halfW_mm, v));
            }
            updatePreview();
        };

        var grpDashEdge = pnl.add('group');
        grpDashEdge.orientation = 'row';
        grpDashEdge.alignChildren = ['fill', 'top'];

        var pnlDash = grpDashEdge.add('panel', undefined, L('panelDash'));
        pnlDash.orientation = 'column';
        pnlDash.alignChildren = ['fill', 'top'];
        pnlDash.margins = [15, 20, 15, 10];

        var grpWidth = pnlDash.add('group');
        grpWidth.add('statictext', [0, 0, labelW2, 20], L('lblLineWidth'));
        var inputWidth = grpWidth.add('edittext', undefined, '3');
        inputWidth.characters = 3;
        changeValueByArrowKey(inputWidth, false, inputWidth);
        grpWidth.add('statictext', undefined, L('unitStroke'));

        var grpGap = pnlDash.add('group');
        grpGap.add('statictext', [0, 0, labelW2, 20], L('lblGap'));
        var inputGap = grpGap.add('edittext', undefined, '6');
        inputGap.characters = 3;
        changeValueByArrowKey(inputGap, false, inputGap);
        grpGap.add('statictext', undefined, L('unitRuler'));

        var grpDashLen = pnlDash.add('group');
        grpDashLen.add('statictext', [0, 0, labelW2, 20], L('lblLength'));
        var inputDashLen = grpDashLen.add('edittext', [0, 0, 40, 18], '0');
        changeValueByArrowKey(inputDashLen, false, inputDashLen);
        grpDashLen.add('statictext', undefined, L('unitRuler'));

        var pnlEdge = grpDashEdge.add('panel', undefined, L('panelEdge'));
        pnlEdge.orientation = 'column';
        pnlEdge.alignChildren = ['left', 'top'];
        pnlEdge.margins = [15, 20, 15, 10];

        var grpEdgeRadio = pnlEdge.add('group');
        grpEdgeRadio.orientation = 'row';
        var rdEdgeNone = grpEdgeRadio.add('radiobutton', undefined, L('rdNone'));
        var rdEdgeCircle = grpEdgeRadio.add('radiobutton', undefined, L('rdCircle'));
        var rdEdgeTriangle = grpEdgeRadio.add('radiobutton', undefined, L('rdTriangle'));
        rdEdgeNone.value = true;

        var grpEdgeSize = pnlEdge.add('group');
        grpEdgeSize.add('statictext', undefined, L('lblSize'));
        var inputEdgeSize = grpEdgeSize.add('edittext', undefined, '5');
        inputEdgeSize.characters = 3;
        changeValueByArrowKey(inputEdgeSize, false, inputEdgeSize);
        grpEdgeSize.add('statictext', undefined, L('unitStroke'));

        var chkEdgeOnly = pnlEdge.add('checkbox', undefined, L('chkShapeOnly'));

        var grpPreviewRow = dlg.add('group');
        grpPreviewRow.orientation = 'row';
        grpPreviewRow.alignment = ['center', 'top'];
        grpPreviewRow.alignChildren = ['left', 'center'];

        var chkPreview = grpPreviewRow.add('checkbox', undefined, L('chkPreview'));
        chkPreview.value = false;

        var chkExpandAppearance = grpPreviewRow.add('checkbox', undefined, L('chkExpandAppearance'));
        chkExpandAppearance.value = false;

        var grpBtn = dlg.add('group');
        grpBtn.alignment = ['center', 'top'];
        grpBtn.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        grpBtn.add('button', undefined, L('btnOK'), { name: 'ok' });

        /* プレビュー更新 / Update preview */
        function updatePreview() {
            removePreviewLayer();
            if (!chkPreview.value) return;
            applyEffect(true);
        }

        /* イベント / Events */
        chkPreview.onClick = function () { updatePreview(); };
        inputWidth.onChanging = updatePreview;
        inputGap.onChanging = updatePreview;
        inputDashLen.onChanging = updatePreview;
        chkLR.onClick = function () { updatePreview(); };
        chkMiddle.onClick = function () {
            inputOffset.enabled = chkMiddle.value;
            lblOffsetMM.enabled = chkMiddle.value;
            sliderOffset.enabled = chkMiddle.value;
            updateEdgeDim();
            updatePreview();
        };
        inputHoleSize.onChanging = updatePreview;
        chkHoleLeft.onClick = function () { updatePreview(); };
        chkHoleRight.onClick = function () { updatePreview(); };
        function updateHoleDim() {
            var on = !rdHoleNone.value;
            grpHoleSide.enabled = on;
            grpHoleSize.enabled = on;
        }
        updateHoleDim();
        rdHoleNone.onClick = function () { updateHoleDim(); updatePreview(); };
        rdHoleCircle.onClick = function () { updateHoleDim(); updatePreview(); };
        rdHoleTriangle.onClick = function () { updateHoleDim(); updatePreview(); };
        function updateEdgeDim() {
            var hasCenter = chkMiddle.value;
            pnlEdge.enabled = hasCenter;

            var on = hasCenter && !rdEdgeNone.value;
            grpEdgeSize.enabled = on;
            chkEdgeOnly.enabled = on;
        }
        updateEdgeDim();
        rdEdgeNone.onClick = function () { updateEdgeDim(); updatePreview(); };
        rdEdgeCircle.onClick = function () { updateEdgeDim(); updatePreview(); };
        rdEdgeTriangle.onClick = function () { updateEdgeDim(); updatePreview(); };
        inputEdgeSize.onChanging = updatePreview;
        chkEdgeOnly.onClick = function () { updatePreview(); };
        function updateCornerDim() {
            grpCornerSize.enabled = !rdCornerNone.value;
        }
        updateCornerDim();
        rdCornerNone.onClick = function () { updateCornerDim(); updatePreview(); };
        rdRound.onClick = function () { updateCornerDim(); updatePreview(); };
        rdInverse.onClick = function () { updateCornerDim(); updatePreview(); };
        rdChamfer.onClick = function () { updateCornerDim(); updatePreview(); };
        inputCornerSize.onChanging = updatePreview;

        if (dlg.show() !== 1) {
            /* キャンセル：プレビューレイヤーを削除して元に戻す / Cancel: remove preview layer and restore originals */
            removePreviewLayer();
            try { previewLayer.remove(); } catch (e) { }
        } else {
            /* OK：プレビューレイヤーを削除し、元オブジェクトに適用 / OK: remove preview layer and apply to originals */
            removePreviewLayer();
            try { previewLayer.remove(); } catch (e) { }

            var strokeW = parseFloat(inputWidth.text);
            var gap = parseFloat(inputGap.text);
            var dashLen = parseFloat(inputDashLen.text);
            var offsetMM = parseFloat(inputOffset.text);
            var edgeSize = parseFloat(inputEdgeSize.text);
            var cornerSize = parseFloat(inputCornerSize.text);
            var holeSize = parseFloat(inputHoleSize.text);

            if (isNaN(strokeW) || isNaN(gap) || isNaN(dashLen) || isNaN(offsetMM) || isNaN(edgeSize) || isNaN(cornerSize) || isNaN(holeSize)) {
                alert(L('alertEnterNumbers'));
            } else {
                doc.activeLayer = originalLayer;
                applyEffect(false);
                if (chkExpandAppearance.value) {
                    app.executeMenuCommand('expandStyle');
                    app.executeMenuCommand('ungroup');
                }
            }
        }
    } finally {
        if (typeof unloadStrokeDotAction === 'function') unloadStrokeDotAction();
        // app.executeMenuCommand('edge');
    }
}
