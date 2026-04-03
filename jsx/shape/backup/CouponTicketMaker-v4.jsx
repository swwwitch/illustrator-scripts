#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * CouponTicketMaker.jsx
 * 概要: 選択した長方形に対して、ミシン目・エッジ・コーナー・ホールを付加し、チケット風の形状を作成します。
 * Summary: Adds perforations, edges, corners, and holes to selected rectangles to create ticket-like shapes.
 * 更新日: 2026-03-08
 * Updated: 2026-03-08
 */

var SCRIPT_VERSION = "v1.1";

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
    panelPerforationLR: { ja: "ミシン目（左右）", en: "Perforation (L/R)" },
    panelPerforationCenter: { ja: "分割線", en: "Divider" },
    panelDash: { ja: "点線", en: "Dotted Line" },
    panelEdge: { ja: "エッジ", en: "Edge" },
    panelCorner: { ja: "コーナー", en: "Corner" },
    chkEnable: { ja: "有効", en: "Enable" },
    lblLineWidth: { ja: "線幅:", en: "Weight:" },
    lblGap: { ja: "間隔:", en: "Gap:" },
    lblLength: { ja: "余白:", en: "Offset:" },
    unitRuler: { ja: rulerUnitLabel, en: rulerUnitLabel },
    unitStroke: { ja: strokeUnitLabel, en: strokeUnitLabel },
    rdNone: { ja: "なし", en: "None" },
    rdDot: { ja: "ドット", en: "Dot" },
    rdDash: { ja: "破線", en: "Dash" },
    rdCircle: { ja: "円", en: "Circle" },
    rdTriangle: { ja: "三角", en: "Triangle" },
    rdRound: { ja: "角丸", en: "Rounded" },
    rdInverse: { ja: "逆角丸", en: "Inverse Round" },
    rdChamfer: { ja: "面取り", en: "Chamfer" },
    panelHole: { ja: "スリット／ホール", en: "Slit / Hole" },
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

    var str = '/version 3 /name [ 9 5374726f6b65446f74 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 3 646f74 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_setStroke) /localizedName [ 12 e7b79ae38292e8a8ade5ae9a ] /isOpen 0 /isOn 1 /hasDialog 0 /parameterCount 12 /parameter-1 { /key 2003072104 /showInPalette 4294967295 /type (unit real) /value 4.0 /unit 592476268 } /parameter-2 { /key 1667330094 /showInPalette 4294967295 /type (enumerated) /name [ 12 e4b8b8e59e8be7b79ae7abaf ] /value 1 } /parameter-3 { /key 1836344690 /showInPalette 4294967295 /type (real) /value 10.0 } /parameter-4 { /key 1785686382 /showInPalette 4294967295 /type (enumerated) /name [ 18 e3839ee382a4e382bfe383bce7b590e59088 ] /value 0 } /parameter-5 { /key 1684825454 /showInPalette 4294967295 /type (integer) /value 2 } /parameter-6 { /key 1685284913 /showInPalette 4294967295 /type (unit real) /value 0.0 /unit 592476268 } /parameter-7 { /key 1685284914 /showInPalette 4294967295 /type (unit real) /value 6.0 /unit 592476268 } /parameter-8 { /key 1684104298 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-9 { /key 1634231345 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-10 { /key 1634231346 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-11 { /key 1634230636 /showInPalette 4294967295 /type (enumerated) /name [ 24 e38391e382b9e381aee7b582e782b9e381abe9858de7bdae ] /value 0 } /parameter-12 { /key 1634494318 /showInPalette 4294967295 /type (enumerated) /name [ 6 e4b8ade5a4ae ] /value 0 } } }';

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

function makeGray60Color(doc) {
    if (doc.documentColorSpace == DocumentColorSpace.CMYK) {
        var c = new CMYKColor();
        c.cyan = 0;
        c.magenta = 0;
        c.yellow = 0;
        c.black = 60;
        return c;
    }
    var g = new GrayColor();
    g.gray = 60;
    return g;
}

function normalizeInitialAppearance(items, doc) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var hasFill = item.filled;
        var hasStroke = item.stroked;

        if (hasFill) {
            item.stroked = false;
            continue;
        }

        if (hasStroke) {
            try {
                item.fillColor = item.strokeColor;
            } catch (e) {
                item.fillColor = makeGray60Color(doc);
            }
            item.filled = true;
            item.stroked = false;
            continue;
        }

        item.fillColor = makeGray60Color(doc);
        item.filled = true;
        item.stroked = false;
    }
}

var doc = app.activeDocument;
var sel = doc.selection;

if (sel.length === 0) {
    alert(L('alertSelectRectangle'));
} else {
    // app.executeMenuCommand('edge');
    try {
        normalizeInitialAppearance(sel, doc);
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

            // ミシン目（左右）の値を取得
            var strokeWValueLR = parseFloat(inputWidthLR.text);
            var gapValueLR = parseFloat(inputGapLR.text);
            var dashLenValueLR = parseFloat(inputDashLenLR.text);
            var strokeWLR = toPt(strokeWValueLR, strokePtFactor);
            var gapLR = toPt(gapValueLR, rulerPtFactor);
            var dashLenLR = toPt(dashLenValueLR, rulerPtFactor);

            // ミシン目（中央）の値を取得
            var strokeWValueC = parseFloat(inputWidthC.text);
            var gapValueC = parseFloat(inputGapC.text);
            var dashLenValueC = parseFloat(inputDashLenC.text);
            var offsetValue = parseFloat(inputOffset.text);
            var strokeWC = toPt(strokeWValueC, strokePtFactor);
            var gapC = toPt(gapValueC, rulerPtFactor);
            var dashLenC = toPt(dashLenValueC, rulerPtFactor);
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

                // ミシン目を作成（中央）
                if (chkCenter.value && !(chkEdgeOnlyC.value && !rdEdgeNoneC.value)) {
                    var line = doc.pathItems.add();
                    line.setEntirePath([[lineX, y - dashLenC], [lineX, y - h + dashLenC]]);
                    line.filled = false;
                    line.stroked = true;

                    doc.selection = null;
                    line.selected = true;
                    act_StrokeDot();
                    line.strokeWidth = strokeWC;
                    if (rdDotC.value) {
                        line.strokeDashes = [0, gapC];
                    } else {
                        line.strokeCap = StrokeCap.BUTTENDCAP;
                        line.strokeDashes = [gapC, gapC];
                    }
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
                        lrLine.setEntirePath([[lrPositions[lr], y - dashLenLR], [lrPositions[lr], y - h + dashLenLR]]);
                        lrLine.filled = false;
                        lrLine.stroked = true;

                        doc.selection = null;
                        lrLine.selected = true;
                        act_StrokeDot();
                        lrLine.strokeWidth = strokeWLR;
                        lrLine.strokeDashes = [0, gapLR];
                        doc.selection = null;
                        lrLine.selected = true;
                        app.executeMenuCommand('Live Outline Stroke');
                        items.push(doc.selection[0]);
                    }
                }

                // エッジを作成（中央）
                if (chkCenter.value && !rdEdgeNoneC.value) {
                    var edgeSizeValueC = parseFloat(inputEdgeSizeC.text);
                    if (!isNaN(edgeSizeValueC) && edgeSizeValueC > 0) {
                        var edgeSizePtC = toPt(edgeSizeValueC, strokePtFactor);
                        var edgeRC = edgeSizePtC / 2;
                        var edgePositionsC = [[lineX, y], [lineX, y - h]];
                        for (var ep = 0; ep < edgePositionsC.length; ep++) {
                            var epx = edgePositionsC[ep][0];
                            var epy = edgePositionsC[ep][1];
                            if (rdEdgeCircleC.value) {
                                var edgeShape = doc.pathItems.ellipse(epy + edgeRC, epx - edgeRC, edgeSizePtC, edgeSizePtC);
                                edgeShape.filled = true;
                                edgeShape.stroked = false;
                            } else {
                                var edgeShape = doc.pathItems.rectangle(epy + edgeRC, epx - edgeRC, edgeSizePtC, edgeSizePtC);
                                edgeShape.filled = true;
                                edgeShape.stroked = false;
                                edgeShape.rotate(45);
                            }
                            items.push(edgeShape);
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
        function removePreviewLayer(skipRedraw) {
            while (previewLayer.pageItems.length > 0) {
                previewLayer.pageItems[0].remove();
            }
            for (var i = 0; i < sel.length; i++) {
                sel[i].hidden = false;
            }
            if (!skipRedraw) app.redraw();
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
        var labelW2 = 78;

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
        var inputHoleSize = grpHoleSize.add('edittext', undefined, '5');
        inputHoleSize.characters = 3;
        changeValueByArrowKey(inputHoleSize, false, inputHoleSize);
        grpHoleSize.add('statictext', undefined, L('unitRuler'));

        // ===== ミシン目（左右）パネル =====
        var pnlLR = grpCornerHole.add('panel', undefined, L('panelPerforationLR'));
        pnlLR.orientation = 'column';
        pnlLR.alignChildren = ['fill', 'top'];
        pnlLR.margins = [15, 20, 15, 10];

        var grpEnableLR = pnlLR.add('group');
        grpEnableLR.margins = [15, 0, 0, 0];
        var chkLR = grpEnableLR.add('checkbox', undefined, L('chkEnable'));
        chkLR.characters = 5;

        var grpWidthLR = pnlLR.add('group');
        grpWidthLR.add('statictext', undefined, L('lblLineWidth'));
        var inputWidthLR = grpWidthLR.add('edittext', undefined, '3');
        inputWidthLR.characters = 3;
        changeValueByArrowKey(inputWidthLR, false, inputWidthLR);
        grpWidthLR.add('statictext', undefined, L('unitStroke'));

        var grpGapLR = pnlLR.add('group');
        grpGapLR.add('statictext', undefined, L('lblGap'));
        var inputGapLR = grpGapLR.add('edittext', undefined, '6');
        inputGapLR.characters = 3;
        changeValueByArrowKey(inputGapLR, false, inputGapLR);
        grpGapLR.add('statictext', undefined, L('unitRuler'));

        var grpDashLenLR = pnlLR.add('group');
        grpDashLenLR.add('statictext', undefined, L('lblLength'));
        var inputDashLenLR = grpDashLenLR.add('edittext', [0, 0, 40, 18], '0');
        changeValueByArrowKey(inputDashLenLR, false, inputDashLenLR);
        grpDashLenLR.add('statictext', undefined, L('unitRuler'));

        // ===== ミシン目（中央）パネル =====
        var pnlCenter = dlg.add('panel', undefined, L('panelPerforationCenter'));
        pnlCenter.orientation = 'column';
        pnlCenter.alignChildren = ['fill', 'top'];
        pnlCenter.margins = [15, 20, 15, 10];

        var grpCenterTop = pnlCenter.add('group');
        grpCenterTop.orientation = 'row';
        grpCenterTop.alignChildren = ['center', 'center'];

        var chkCenter = grpCenterTop.add('checkbox', undefined, L('chkEnable'));
        chkCenter.characters = 5;
        chkCenter.value = true;
        var inputOffset = grpCenterTop.add('edittext', undefined, '0');
        inputOffset.characters = 5;
        changeValueByArrowKey(inputOffset, true, inputOffset);
        var lblOffsetMM = grpCenterTop.add('statictext', undefined, L('unitRuler'));

        var halfW_mm = Math.round(fromPt(sel[0].width / 2, rulerPtFactor) * 10) / 10;
        var sliderOffset = grpCenterTop.add('slider', undefined, 0, -halfW_mm, halfW_mm);
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

        var grpDashEdgeC = pnlCenter.add('group');
        grpDashEdgeC.orientation = 'row';
        grpDashEdgeC.alignChildren = ['fill', 'top'];

        var pnlDashC = grpDashEdgeC.add('panel', undefined, L('panelDash'));
        pnlDashC.orientation = 'column';
        pnlDashC.alignChildren = ['fill', 'top'];
        pnlDashC.margins = [15, 20, 15, 10];

        var grpDashRadioC = pnlDashC.add('group');
        grpDashRadioC.orientation = 'row';
        var rdDotC = grpDashRadioC.add('radiobutton', undefined, L('rdDot'));
        var rdDashC = grpDashRadioC.add('radiobutton', undefined, L('rdDash'));
        rdDotC.value = true;

        var grpWidthC = pnlDashC.add('group');
        grpWidthC.add('statictext', undefined, L('lblLineWidth'));
        var inputWidthC = grpWidthC.add('edittext', undefined, '3');
        inputWidthC.characters = 3;
        changeValueByArrowKey(inputWidthC, false, inputWidthC);
        grpWidthC.add('statictext', undefined, L('unitStroke'));

        var grpGapC = pnlDashC.add('group');
        grpGapC.add('statictext', undefined, L('lblGap'));
        var inputGapC = grpGapC.add('edittext', undefined, '6');
        inputGapC.characters = 3;
        changeValueByArrowKey(inputGapC, false, inputGapC);
        grpGapC.add('statictext', undefined, L('unitRuler'));

        var grpDashLenC = pnlDashC.add('group');
        grpDashLenC.add('statictext', undefined, L('lblLength'));
        var inputDashLenC = grpDashLenC.add('edittext', [0, 0, 40, 18], '0');
        changeValueByArrowKey(inputDashLenC, false, inputDashLenC);
        grpDashLenC.add('statictext', undefined, L('unitRuler'));

        var pnlEdgeC = grpDashEdgeC.add('panel', undefined, L('panelEdge'));
        pnlEdgeC.orientation = 'column';
        pnlEdgeC.alignChildren = ['left', 'top'];
        pnlEdgeC.margins = [15, 20, 15, 10];

        var grpEdgeRadioC = pnlEdgeC.add('group');
        grpEdgeRadioC.orientation = 'row';
        var rdEdgeNoneC = grpEdgeRadioC.add('radiobutton', undefined, L('rdNone'));
        var rdEdgeCircleC = grpEdgeRadioC.add('radiobutton', undefined, L('rdCircle'));
        var rdEdgeTriangleC = grpEdgeRadioC.add('radiobutton', undefined, L('rdTriangle'));
        rdEdgeNoneC.value = true;

        var grpEdgeSizeC = pnlEdgeC.add('group');
        grpEdgeSizeC.add('statictext', undefined, L('lblSize'));
        var inputEdgeSizeC = grpEdgeSizeC.add('edittext', undefined, '5');
        inputEdgeSizeC.characters = 3;
        changeValueByArrowKey(inputEdgeSizeC, false, inputEdgeSizeC);
        grpEdgeSizeC.add('statictext', undefined, L('unitStroke'));

        var chkEdgeOnlyC = pnlEdgeC.add('checkbox', undefined, L('chkShapeOnly'));

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
            removePreviewLayer(chkPreview.value);
            if (!chkPreview.value) return;
            applyEffect(true);
        }

        /* イベント / Events */
        chkPreview.onClick = function () { updatePreview(); };

        // ミシン目（左右）イベント
        chkLR.onClick = function () {
            updateLRDim();
            updatePreview();
        };
        inputWidthLR.onChanging = updatePreview;
        inputGapLR.onChanging = updatePreview;
        inputDashLenLR.onChanging = updatePreview;
        function updateLRDim() {
            var on = chkLR.value;
            grpWidthLR.enabled = on;
            grpGapLR.enabled = on;
            grpDashLenLR.enabled = on;
        }
        updateLRDim();

        // ミシン目（中央）イベント
        chkCenter.onClick = function () {
            updateCenterDim();
            updatePreview();
        };
        rdDotC.onClick = function () { updatePreview(); };
        rdDashC.onClick = function () { updatePreview(); };
        inputWidthC.onChanging = updatePreview;
        inputGapC.onChanging = updatePreview;
        inputDashLenC.onChanging = updatePreview;
        function updateCenterDim() {
            var on = chkCenter.value;
            inputOffset.enabled = on;
            lblOffsetMM.enabled = on;
            sliderOffset.enabled = on;
            grpDashEdgeC.enabled = on;
        }
        function syncDashLenFromEdgeSizeC() {
            if (!chkCenter.value || rdEdgeNoneC.value) return;

            var edgeSizeValue = parseFloat(inputEdgeSizeC.text);
            if (isNaN(edgeSizeValue)) return;

            var dashLenValue = edgeSizeValue;
            if (rdEdgeCircleC.value) {
                dashLenValue = edgeSizeValue * 0.75;
            }

            dashLenValue = Math.round(dashLenValue * 1000) / 1000;
            inputDashLenC.text = String(dashLenValue);
        }

        function updateEdgeDimC() {
            var on = chkCenter.value && !rdEdgeNoneC.value;
            grpEdgeSizeC.enabled = on;
            chkEdgeOnlyC.enabled = on;
            pnlDashC.enabled = chkCenter.value && !(on && chkEdgeOnlyC.value);
            if (on) syncDashLenFromEdgeSizeC();
        }
        updateCenterDim();
        updateEdgeDimC();
        rdEdgeNoneC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        rdEdgeCircleC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        rdEdgeTriangleC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        inputEdgeSizeC.onChanging = function () { syncDashLenFromEdgeSizeC(); updatePreview(); };
        chkEdgeOnlyC.onClick = function () { updateEdgeDimC(); updatePreview(); };

        // ホール・コーナーイベント
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
        function updateCornerDim() {
            grpCornerSize.enabled = !rdCornerNone.value;
        }
        updateCornerDim();
        rdCornerNone.onClick = function () { updateCornerDim(); updatePreview(); };
        rdRound.onClick = function () { updateCornerDim(); updatePreview(); };
        rdInverse.onClick = function () { updateCornerDim(); updatePreview(); };
        rdChamfer.onClick = function () { updateCornerDim(); updatePreview(); };
        inputCornerSize.onChanging = updatePreview;

        syncDashLenFromEdgeSizeC();
        if (dlg.show() !== 1) {
            /* キャンセル：プレビューレイヤーを削除して元に戻す / Cancel: remove preview layer and restore originals */
            removePreviewLayer();
            try { previewLayer.remove(); } catch (e) { }
        } else {
            /* OK：プレビューレイヤーを削除し、元オブジェクトに適用 / OK: remove preview layer and apply to originals */
            removePreviewLayer();
            try { previewLayer.remove(); } catch (e) { }

            var cornerSize = parseFloat(inputCornerSize.text);
            var holeSize = parseFloat(inputHoleSize.text);

            if (isNaN(cornerSize) || isNaN(holeSize)) {
                alert(L('alertEnterNumbers'));
            } else {
                doc.activeLayer = originalLayer;
                applyEffect(false);
                if (chkExpandAppearance.value) {
                    app.executeMenuCommand('expandStyle');
                    // app.executeMenuCommand('ungroup');
                }
            }
        }
    } finally {
        if (typeof unloadStrokeDotAction === 'function') unloadStrokeDotAction();
        // app.executeMenuCommand('edge');
    }
}
