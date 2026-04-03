#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * CouponTicketMaker.jsx
 * 概要:
 * 選択した長方形パスを元に、ミシン目（左右・中央）、ギザギザ、エッジ、コーナー、スリット／ホールなどを
 * 組み合わせてチケット風の形状を生成するスクリプトです。
 * プレビュー機能によりダイアログ上で結果を確認しながら調整できます。
 * プレビューは専用レイヤーで生成され、確定時のみ元オブジェクトへ適用されます。
 *
 * Summary:
 * Generates ticket-style shapes from a selected rectangle by combining perforations
 * (left/right and center), zigzag edges, corner processing, and slit/hole shapes.
 * A live preview allows adjusting parameters before applying the result.
 * Preview objects are created on a temporary layer and applied to the original
 * object only when confirmed.
 *
 * 更新日: 2026-03-14
 * Updated: 2026-03-14
 */

var SCRIPT_VERSION = "v1.2";

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

// strokeUnits (used only for divider line width)
var strokeUnitCode = app.preferences.getIntegerPreference("strokeUnits");
var strokeUnitLabel = getUnitLabel(strokeUnitCode, "strokeUnits");
var strokePtFactor = getPtFactorFromUnitCode(strokeUnitCode);

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "チケットメーカー", en: "Ticket Maker" },
    alertSelectRectangle: { ja: "長方形を選択してください。", en: "Please select a rectangle." },
    alertOpenDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    alertRectangleOnly: { ja: "長方形のパスを1つだけ選択してください。", en: "Please select exactly one rectangular path." },
    alertSingleOnly: { ja: "複数選択時は実行できません。オブジェクトを1つだけ選択してください。", en: "This script cannot run with multiple selections. Please select only one object." },
    alertGroupNotAllowed: { ja: "グループを選択しているときは実行できません。単体のオブジェクトを選択してください。", en: "This script cannot run when a group is selected. Please select a single object." },
    alertEnterNumbers: { ja: "数値を入力してください。", en: "Please enter numeric values." },
    alertEnterValidNumbers: { ja: "数値欄に正しい数値を入力してください。", en: "Please enter valid numeric values in the numeric fields." },
    previewLayerName: { ja: "プレビュー", en: "Preview" },
    panelLR: { ja: "左右", en: "L/R" },
    panelPerforationLR: { ja: "ミシン目", en: "Perforation" },
    panelZigzag: { ja: "ギザギザ", en: "Zigzag" },
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
    rdLR: { ja: "左右", en: "L/R" },
    rdTB: { ja: "上下", en: "T/B" },
    lblZZSize: { ja: "大きさ:", en: "Size:" },
    lblZZRepeat: { ja: "繰り返し:", en: "Repeat:" },
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

function saveSelection(doc) {
    var result = [];
    try {
        var sel = doc.selection;
        for (var i = 0; i < sel.length; i++) result.push(sel[i]);
    } catch (e) {
    }
    return result;
}

function restoreSelection(doc, items) {
    doc.selection = null;
    if (!items) return;
    for (var i = 0; i < items.length; i++) {
        try {
            items[i].selected = true;
        } catch (e) {
        }
    }
}

function selectItems(doc, items) {
    doc.selection = null;
    for (var i = 0; i < items.length; i++) {
        items[i].selected = true;
    }
}

function outlineStrokeItem(doc, item) {
    var prevSelection = saveSelection(doc);
    try {
        selectItems(doc, [item]);
        app.executeMenuCommand('Live Outline Stroke');
        if (!doc.selection || doc.selection.length === 0) {
            throw new Error('Live Outline Stroke failed.');
        }
        return doc.selection[0];
    } finally {
        restoreSelection(doc, prevSelection);
    }
}

function groupItems(doc, items) {
    if (!items || items.length === 0) return null;
    if (items.length === 1) return items[0];

    var prevSelection = saveSelection(doc);
    try {
        selectItems(doc, items);
        app.executeMenuCommand('group');
        if (!doc.selection || doc.selection.length === 0) {
            throw new Error('Group command failed.');
        }
        return doc.selection[0];
    } finally {
        restoreSelection(doc, prevSelection);
    }
}

function getSingleSelection(doc) {
    if (!doc.selection || doc.selection.length === 0) return null;
    return doc.selection[0];
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

function isRectanglePath(item) {
    if (!item || item.typename !== 'PathItem') return false;
    if (item.guides || item.clipping) return false;
    if (item.pathPoints.length !== 4) return false;

    var gb = item.geometricBounds;
    var left = gb[0];
    var top = gb[1];
    var right = gb[2];
    var bottom = gb[3];
    var tol = 0.01;

    function isNear(a, b) {
        return Math.abs(a - b) <= tol;
    }

    var pts = [];
    for (var i = 0; i < 4; i++) {
        var p = item.pathPoints[i].anchor;
        pts.push([p[0], p[1]]);
    }

    var hasLT = false, hasRT = false, hasRB = false, hasLB = false;
    for (var j = 0; j < pts.length; j++) {
        var x = pts[j][0];
        var y = pts[j][1];
        if (isNear(x, left) && isNear(y, top)) hasLT = true;
        else if (isNear(x, right) && isNear(y, top)) hasRT = true;
        else if (isNear(x, right) && isNear(y, bottom)) hasRB = true;
        else if (isNear(x, left) && isNear(y, bottom)) hasLB = true;
        else return false;
    }

    return hasLT && hasRT && hasRB && hasLB;
}

function getSafeSelection(doc) {
    try {
        return doc.selection;
    } catch (e) {
        return [];
    }
}

function validateStartupState() {
    if (app.documents.length === 0) {
        alert(L('alertOpenDocument'));
        return null;
    }

    var doc = app.activeDocument;
    var sel = getSafeSelection(doc);

    if (sel.length === 0) {
        alert(L('alertSelectRectangle'));
        return null;
    }
    if (sel.length > 1) {
        alert(L('alertSingleOnly'));
        return null;
    }
    if (sel[0].typename === 'GroupItem') {
        alert(L('alertGroupNotAllowed'));
        return null;
    }
    if (!isRectanglePath(sel[0])) {
        alert(L('alertRectangleOnly'));
        return null;
    }

    return {
        doc: doc,
        sel: sel
    };
}

function main() {
    var startup = validateStartupState();
    if (!startup) return;

    var doc = startup.doc;
    var sel = startup.sel;
    // app.executeMenuCommand('edge');
    try {
        var originalLayer = sel[0].layer;

        // プレビュー専用レイヤーは遅延作成
        var previewLayer = null;

        function hasUsableLayer(layer) {
            if (!layer) return false;
            try {
                var _ = layer.name;
                return true;
            } catch (e) {
                return false;
            }
        }

        function ensurePreviewLayer() {
            if (hasUsableLayer(previewLayer)) return previewLayer;
            previewLayer = doc.layers.add();
            previewLayer.name = L('previewLayerName');
            return previewLayer;
        }

        function collectUIValues() {
            return {
                // L/R perforation
                strokeWValueLR: parseFloat(inputWidthLR.text),
                gapValueLR: parseFloat(inputGapLR.text),
                dashLenValueLR: parseFloat(inputDashLenLR.text),

                // zigzag
                zzSizeValue: parseFloat(inputWidthZZ.text),
                zzGapValue: parseFloat(inputGapZZ.text) || 0,
                zzRepeatValue: Math.max(1, Math.round(parseFloat(inputDashLenZZ.text) || 1)),

                // center perforation
                strokeWValueC: parseFloat(inputWidthC.text),
                gapValueC: parseFloat(inputGapC.text),
                dashLenValueC: parseFloat(inputDashLenC.text),
                offsetValue: parseFloat(inputOffset.text),

                // corner
                cornerSizeValue: parseFloat(inputCornerSize.text),

                // hole
                holeSizeValue: parseFloat(inputHoleSize.text),

                // edge
                edgeSizeValueC: parseFloat(inputEdgeSizeC.text)
            };
        }

        function validateUIValues(ui) {
            var numericKeys = [
                'strokeWValueLR', 'gapValueLR', 'dashLenValueLR',
                'zzSizeValue', 'zzGapValue', 'zzRepeatValue',
                'strokeWValueC', 'gapValueC', 'dashLenValueC', 'offsetValue',
                'cornerSizeValue', 'holeSizeValue', 'edgeSizeValueC'
            ];
            for (var i = 0; i < numericKeys.length; i++) {
                var key = numericKeys[i];
                if (isNaN(ui[key])) return false;
            }
            return true;
        }

        function applyRoundCorners(targets, cornerSizeValue) {
            if (!rdRound.value) return;
            if (isNaN(cornerSizeValue) || cornerSizeValue <= 0) return;

            var cornerSizePt = toPt(cornerSizeValue, rulerPtFactor);
            var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSizePt + ' "/></LiveEffect>';
            for (var i = 0; i < targets.length; i++) {
                targets[i].applyEffect(xml);
            }
        }

        function addCenterPerforation(doc, items, x, y, w, h, lineX, dashLenC, strokeWC, gapC) {
            if (chkCenter.value && !(chkEdgeOnlyC.value && !rdEdgeNoneC.value)) {
                var line = doc.pathItems.add();
                line.setEntirePath([[lineX, y - dashLenC], [lineX, y - h + dashLenC]]);
                line.filled = false;
                line.stroked = true;

                var prevSelectionCenter = saveSelection(doc);
                try {
                    selectItems(doc, [line]);
                    act_StrokeDot();
                } finally {
                    restoreSelection(doc, prevSelectionCenter);
                }
                line.strokeWidth = strokeWC;
                if (rdDotC.value) {
                    line.strokeDashes = [0, gapC];
                } else {
                    line.strokeCap = StrokeCap.BUTTENDCAP;
                    line.strokeDashes = [gapC, gapC];
                }
                items.push(outlineStrokeItem(doc, line));
            }
        }

        function addLRPerforation(doc, items, x, y, w, h, dashLenLR, strokeWLR, gapLR) {
            if (chkLR.value) {
                var lrPositions = [x, x + w];
                for (var lr = 0; lr < lrPositions.length; lr++) {
                    var lrLine = doc.pathItems.add();
                    lrLine.setEntirePath([[lrPositions[lr], y - dashLenLR], [lrPositions[lr], y - h + dashLenLR]]);
                    lrLine.filled = false;
                    lrLine.stroked = true;

                    var prevSelectionLR = saveSelection(doc);
                    try {
                        selectItems(doc, [lrLine]);
                        act_StrokeDot();
                    } finally {
                        restoreSelection(doc, prevSelectionLR);
                    }
                    lrLine.strokeWidth = strokeWLR;
                    lrLine.strokeDashes = [0, gapLR];
                    items.push(outlineStrokeItem(doc, lrLine));
                }
            }
        }

        function addZigzag(doc, items, x, y, w, h, zzSizePt, zzGapPt, zzRepeat) {
            if (!rdZZNone.value && zzSizePt > 0) {
                // zzSizePtは対角線の長さ、正方形の辺に変換
                var zzSide = zzSizePt / Math.sqrt(2);
                var halfD = zzSizePt / 2;
                var zigzagItems = [];

                var step = zzSizePt + zzGapPt;
                var total = step * zzRepeat - zzGapPt;

                if (rdZZLR.value) {
                    // 左右：辺の垂直中央を基準に配置
                    var centerY = y - h / 2;
                    var startY = centerY + total / 2 - halfD;
                    var lrXPositions = [x, x + w];
                    for (var lr = 0; lr < lrXPositions.length; lr++) {
                        var lrX = lrXPositions[lr];
                        for (var di = 0; di < zzRepeat; di++) {
                            var dy = startY - step * di;
                            var diamond = doc.pathItems.rectangle(dy + zzSide / 2, lrX - zzSide / 2, zzSide, zzSide);
                            diamond.filled = true;
                            diamond.stroked = false;
                            diamond.rotate(45);
                            zigzagItems.push(diamond);
                        }
                    }
                } else {
                    // 上下：辺の水平中央を基準に配置
                    var centerX = x + w / 2;
                    var startX = centerX - total / 2 + halfD;
                    var tbYPositions = [y, y - h];
                    for (var tb = 0; tb < tbYPositions.length; tb++) {
                        var tbY = tbYPositions[tb];
                        for (var di = 0; di < zzRepeat; di++) {
                            var dx = startX + step * di;
                            var diamond = doc.pathItems.rectangle(tbY + zzSide / 2, dx - zzSide / 2, zzSide, zzSide);
                            diamond.filled = true;
                            diamond.stroked = false;
                            diamond.rotate(45);
                            zigzagItems.push(diamond);
                        }
                    }
                }

                if (zigzagItems.length > 0) {
                    items.push(groupItems(doc, zigzagItems));
                }
            }
        }

        function addHoles(doc, items, x, y, w, h, holeSizeValue) {
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
                    items.push(groupItems(doc, holeItems));
                }
            }
        }

        function addInverseCorners(doc, items, x, y, w, h, cornerSizeValue) {
            if (rdInverse.value) {
                if (!isNaN(cornerSizeValue) && cornerSizeValue > 0) {
                    var cSizePt = toPt(cornerSizeValue, rulerPtFactor);

                    // 上辺（左上・右上の角）
                    var topLine = doc.pathItems.add();
                    topLine.setEntirePath([[x, y], [x + w, y]]);
                    topLine.filled = false;
                    topLine.stroked = true;

                    var prevSelectionTop = saveSelection(doc);
                    try {
                        selectItems(doc, [topLine]);
                        act_StrokeDot();
                    } finally {
                        restoreSelection(doc, prevSelectionTop);
                    }
                    topLine.strokeWidth = cSizePt;
                    topLine.strokeDashes = [0, 1000];
                    items.push(outlineStrokeItem(doc, topLine));

                    // 下辺（左下・右下の角）
                    var bottomLine = doc.pathItems.add();
                    bottomLine.setEntirePath([[x, y - h], [x + w, y - h]]);
                    bottomLine.filled = false;
                    bottomLine.stroked = true;

                    var prevSelectionBottom = saveSelection(doc);
                    try {
                        selectItems(doc, [bottomLine]);
                        act_StrokeDot();
                    } finally {
                        restoreSelection(doc, prevSelectionBottom);
                    }
                    bottomLine.strokeWidth = cSizePt;
                    bottomLine.strokeDashes = [0, 1000];
                    items.push(outlineStrokeItem(doc, bottomLine));
                }
            }
        }

        function addChamferCorners(doc, items, x, y, w, h, cornerSizeValue) {
            if (rdChamfer.value) {
                if (!isNaN(cornerSizeValue) && cornerSizeValue > 0) {
                    var cSizePt = toPt(cornerSizeValue, rulerPtFactor);
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
        }

        function addCenterEdges(doc, items, lineX, y, h, edgeSizeValueC) {
            if (chkCenter.value && !rdEdgeNoneC.value) {
                if (!isNaN(edgeSizeValueC) && edgeSizeValueC > 0) {
                    var edgeSizePtC = toPt(edgeSizeValueC, rulerPtFactor);
                    var edgeRC = edgeSizePtC / 2;
                    var edgePositionsC = [[lineX, y], [lineX, y - h]];
                    for (var ep = 0; ep < edgePositionsC.length; ep++) {
                        var epx = edgePositionsC[ep][0];
                        var epy = edgePositionsC[ep][1];
                        var edgeShape;
                        if (rdEdgeCircleC.value) {
                            edgeShape = doc.pathItems.ellipse(epy + edgeRC, epx - edgeRC, edgeSizePtC, edgeSizePtC);
                            edgeShape.filled = true;
                            edgeShape.stroked = false;
                        } else {
                            edgeShape = doc.pathItems.rectangle(epy + edgeRC, epx - edgeRC, edgeSizePtC, edgeSizePtC);
                            edgeShape.filled = true;
                            edgeShape.stroked = false;
                            edgeShape.rotate(45);
                        }
                        items.push(edgeShape);
                    }
                }
            }
        }

        function finalizeSubtract(doc, items, finalResults) {
            var groupedItems = groupItems(doc, items);
            selectItems(doc, [groupedItems]);
            app.executeMenuCommand('Live Pathfinder Subtract');
            var subtractResult = getSingleSelection(doc);
            if (subtractResult) {
                finalResults.push(subtractResult);
            }
        }
        /* エフェクトを適用する関数 / Apply effect */
        /* isPreview=true: プレビューレイヤーに複製して適用、元を非表示 / Duplicate to preview layer and hide originals */
        /* isPreview=false: 元のオブジェクトに直接適用 / Apply directly to original objects */
        function applyEffect(isPreview) {
            var targets = [];
            var finalResults = [];

            for (var i = 0; i < sel.length; i++) {
                if (isPreview) {
                    var targetLayer = ensurePreviewLayer();
                    doc.activeLayer = targetLayer;
                    var target = sel[i].duplicate(targetLayer, ElementPlacement.PLACEATEND);
                    normalizeInitialAppearance([target], doc);
                    sel[i].hidden = true;
                    targets.push(target);
                } else {
                    normalizeInitialAppearance([sel[i]], doc);
                    targets.push(sel[i]);
                }
            }

            var ui = collectUIValues();
            if (!validateUIValues(ui)) {
                throw new Error(L('alertEnterValidNumbers'));
            }

            // ミシン目（左右）
            var strokeWLR = toPt(ui.strokeWValueLR, rulerPtFactor);
            var gapLR = toPt(ui.gapValueLR, rulerPtFactor);
            var dashLenLR = toPt(ui.dashLenValueLR, rulerPtFactor);

            // ギザギザ
            var zzSizePt = toPt(ui.zzSizeValue, rulerPtFactor);
            var zzGapPt = toPt(ui.zzGapValue, rulerPtFactor);
            var zzRepeat = ui.zzRepeatValue;

            // ミシン目（中央）
            var strokeWC = toPt(ui.strokeWValueC, strokePtFactor);
            var gapC = toPt(ui.gapValueC, rulerPtFactor);
            var dashLenC = toPt(ui.dashLenValueC, rulerPtFactor);
            var offsetPt = toPt(ui.offsetValue, rulerPtFactor);

            // 角丸を適用
            applyRoundCorners(targets, ui.cornerSizeValue);

            for (var i = 0; i < targets.length; i++) {
                var rect = targets[i];
                var x = rect.left;
                var y = rect.top;
                var w = rect.width;
                var h = rect.height;
                var lineX = x + w / 2 + offsetPt;
                var items = [rect];

                // ミシン目を作成（中央）
                addCenterPerforation(doc, items, x, y, w, h, lineX, dashLenC, strokeWC, gapC);

                // ミシン目を作成（左右）
                addLRPerforation(doc, items, x, y, w, h, dashLenLR, strokeWLR, gapLR);

                // ギザギザを作成
                addZigzag(doc, items, x, y, w, h, zzSizePt, zzGapPt, zzRepeat);

                // エッジを作成（中央）
                addCenterEdges(doc, items, lineX, y, h, ui.edgeSizeValueC);

                // ホールを作成
                addHoles(doc, items, x, y, w, h, ui.holeSizeValue);

                // コーナーの逆角丸を作成
                addInverseCorners(doc, items, x, y, w, h, ui.cornerSizeValue);

                // コーナーの面取りを作成
                addChamferCorners(doc, items, x, y, w, h, ui.cornerSizeValue);

                // グループ化してパスファインダー適用
                finalizeSubtract(doc, items, finalResults);
            }

            if (!isPreview) {
                selectItems(doc, finalResults.length > 0 ? finalResults : targets);
            }
            app.redraw();
        }

        /* プレビューレイヤーをクリアし元オブジェクトを復元 / Clear preview layer and restore original objects */
        function removePreviewLayer(skipRedraw) {
            if (hasUsableLayer(previewLayer)) {
                try {
                    while (previewLayer.pageItems.length > 0) {
                        previewLayer.pageItems[0].remove();
                    }
                } catch (e) {
                }
            }
            for (var i = 0; i < sel.length; i++) {
                sel[i].hidden = false;
            }
            if (!skipRedraw) app.redraw();
        }

        function removePreviewLayerContainer() {
            if (!hasUsableLayer(previewLayer)) return;
            try {
                while (previewLayer.pageItems.length > 0) {
                    previewLayer.pageItems[0].remove();
                }
            } catch (e) {
            }
            try {
                previewLayer.remove();
            } catch (e) {
            }
            previewLayer = null;
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

        // ===== 1行目: コーナー／ギザギザ =====
        var grpRow1 = dlg.add('group');
        grpRow1.orientation = 'row';
        grpRow1.alignChildren = ['fill', 'top'];

        var pnlCorner = grpRow1.add('panel', undefined, L('panelCorner'));
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

        // ===== ギザギザパネル =====
        var pnlZigzag = grpRow1.add('panel', undefined, L('panelZigzag'));
        pnlZigzag.orientation = 'column';
        pnlZigzag.alignChildren = ['fill', 'top'];
        pnlZigzag.margins = [15, 20, 15, 10];

        var grpDirZZ = pnlZigzag.add('group');
        grpDirZZ.orientation = 'row';
        var rdZZNone = grpDirZZ.add('radiobutton', undefined, L('rdNone'));
        var rdZZLR = grpDirZZ.add('radiobutton', undefined, L('rdLR'));
        var rdZZTB = grpDirZZ.add('radiobutton', undefined, L('rdTB'));
        rdZZNone.value = true;

        var zzLabelW = (lang === 'ja') ? 60 : 55;

        var grpWidthZZ = pnlZigzag.add('group');
        var lblWidthZZ = grpWidthZZ.add('statictext', undefined, L('lblZZSize'));
        lblWidthZZ.preferredSize.width = zzLabelW;
        var inputWidthZZ = grpWidthZZ.add('edittext', undefined, '10');
        inputWidthZZ.characters = 3;
        changeValueByArrowKey(inputWidthZZ, false, inputWidthZZ);
        grpWidthZZ.add('statictext', undefined, L('unitStroke'));

        var grpDashLenZZ = pnlZigzag.add('group');
        var lblRepeatZZ = grpDashLenZZ.add('statictext', undefined, L('lblZZRepeat'));
        lblRepeatZZ.preferredSize.width = zzLabelW;
        var inputDashLenZZ = grpDashLenZZ.add('edittext', undefined, '3');
        inputDashLenZZ.characters = 3;
        changeValueByArrowKey(inputDashLenZZ, false, inputDashLenZZ);

        var grpGapZZ = pnlZigzag.add('group');
        var lblGapZZ = grpGapZZ.add('statictext', undefined, L('lblGap'));
        lblGapZZ.preferredSize.width = zzLabelW;
        var inputGapZZ = grpGapZZ.add('edittext', undefined, '0');
        inputGapZZ.characters = 3;
        changeValueByArrowKey(inputGapZZ, true, inputGapZZ);
        grpGapZZ.add('statictext', undefined, L('unitStroke'));

        // ===== 2行目: ミシン目／スリット・ホール =====
        var grpRow2 = dlg.add('group');
        grpRow2.orientation = 'row';
        grpRow2.alignChildren = ['fill', 'top'];

        // ===== 左右パネル =====
        var pnlLRWrap = grpRow2.add('panel', undefined, L('panelLR'));
        pnlLRWrap.orientation = 'row';
        pnlLRWrap.alignChildren = ['fill', 'top'];
        pnlLRWrap.margins = [15, 20, 15, 15];

        // ===== ミシン目パネル =====
        var pnlLR = pnlLRWrap.add('panel', undefined, L('panelPerforationLR'));
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
        grpWidthLR.add('statictext', undefined, L('unitRuler'));

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

        // ===== スリット／ホールパネル =====
        var pnlHole = pnlLRWrap.add('panel', undefined, L('panelHole'));
        pnlHole.orientation = 'column';
        pnlHole.alignChildren = ['left', 'top'];
        pnlHole.margins = [15, 20, 15, 10];

        var grpHoleRadio = pnlHole.add('group');
        grpHoleRadio.orientation = 'row';
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
        var inputHoleSize = grpHoleSize.add('edittext', undefined, '10');
        inputHoleSize.characters = 3;
        changeValueByArrowKey(inputHoleSize, false, inputHoleSize);
        grpHoleSize.add('statictext', undefined, L('unitRuler'));

        // ===== 3行目: 分割線 =====
        // ===== ミシン目（中央）パネル =====
        var pnlCenter = dlg.add('panel', undefined, L('panelPerforationCenter'));
        pnlCenter.orientation = 'column';
        pnlCenter.alignChildren = ['fill', 'top'];
        pnlCenter.margins = [15, 20, 15, 10];

        var grpCenterTop = pnlCenter.add('group');
        grpCenterTop.orientation = 'row';
        grpCenterTop.alignChildren = ['center', 'center'];
        grpCenterTop.margins = [0, 0, 0, 10];

        var chkCenter = grpCenterTop.add('checkbox', undefined, L('chkEnable'));
        chkCenter.characters = 5;
        chkCenter.value = true;
        var inputOffset = grpCenterTop.add('edittext', undefined, '0');
        inputOffset.characters = 5;
        changeValueByArrowKey(inputOffset, true, inputOffset);
        var lblOffsetMM = grpCenterTop.add('statictext', undefined, L('unitRuler'));

        var minHalfW = null;
        for (var si = 0; si < sel.length; si++) {
            var itemHalfW = Math.abs(sel[si].width) / 2;
            if (minHalfW === null || itemHalfW < minHalfW) {
                minHalfW = itemHalfW;
            }
        }
        var halfW_mm = Math.round(fromPt((minHalfW || 0), rulerPtFactor) * 10) / 10;
        // 複数選択時は最小幅の半分を上限にして、すべての対象で安全な範囲に制限
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
        var inputEdgeSizeC = grpEdgeSizeC.add('edittext', undefined, '10');
        inputEdgeSizeC.characters = 3;
        changeValueByArrowKey(inputEdgeSizeC, false, inputEdgeSizeC);
        grpEdgeSizeC.add('statictext', undefined, L('unitStroke'));

        var chkEdgeOnlyC = pnlEdgeC.add('checkbox', undefined, L('chkShapeOnly'));

        var grpPreviewRow = dlg.add('group');
        grpPreviewRow.orientation = 'row';
        grpPreviewRow.alignment = ['center', 'top'];
        grpPreviewRow.alignChildren = ['left', 'center'];

        var chkPreview = grpPreviewRow.add('checkbox', undefined, L('chkPreview'));
        chkPreview.value = true;

        var chkExpandAppearance = grpPreviewRow.add('checkbox', undefined, L('chkExpandAppearance'));
        chkExpandAppearance.value = false;

        var grpBtn = dlg.add('group');
        grpBtn.alignment = ['center', 'top'];
        grpBtn.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        grpBtn.add('button', undefined, L('btnOK'), { name: 'ok' });

        /* プレビュー更新 / Update preview */
        function updatePreview() {
            updateZZDim();
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

        // ギザギザイベント
        rdZZNone.onClick = function () { updateZZDim(); updatePreview(); };
        rdZZLR.onClick = function () { updateZZDim(); updatePreview(); };
        rdZZTB.onClick = function () { updateZZDim(); updatePreview(); };
        inputWidthZZ.onChanging = updatePreview;
        inputGapZZ.onChanging = updatePreview;
        inputDashLenZZ.onChanging = function () { updateZZDim(); updatePreview(); };
        function updateZZDim() {
            var on = !rdZZNone.value;
            grpWidthZZ.enabled = on;
            grpDashLenZZ.enabled = on;
            var repeat = Math.round(parseFloat(inputDashLenZZ.text) || 0);
            grpGapZZ.enabled = on && repeat > 1;
        }
        updateZZDim();

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

            var dashLenValue = edgeSizeValue * 1.2;
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
            removePreviewLayerContainer();
        } else {
            /* OK：プレビューレイヤーを削除し、元オブジェクトに適用 / OK: remove preview layer and apply to originals */
            removePreviewLayer();
            removePreviewLayerContainer();

            var finalUI = collectUIValues();
            if (!validateUIValues(finalUI)) {
                alert(L('alertEnterValidNumbers'));
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

main();