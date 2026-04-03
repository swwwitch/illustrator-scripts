#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * CouponTicketMaker.jsx
 * 概要:
 * 選択した長方形パスを元に、ミシン目（左右・中央）、ギザギザ、エッジ、コーナー、スリット／ホールなどを
 * 組み合わせてチケット風の形状を生成するスクリプトです。
 * ダイアログ最上部にはプリセット選択用のプルダウンメニューと［保存］ボタンを備え、
 * セッション中に現在設定をプリセットとして追加保存できます。
 * 保存した内容はコード組み込み用の形式でも確認でき、プレビュー機能により結果を確認しながら調整できます。
 * プレビューは専用レイヤーで生成され、確定時のみ元オブジェクトへ適用されます。
 * 
 * 紹介記事（note）
 * https://note.com/dtp_tranist/n/n2e949946228a
 *
 * Summary:
 * Generates ticket-style shapes from a selected rectangle by combining perforations
 * (left/right and center), zigzag edges, corner processing, and slit/hole shapes.
 * The dialog now includes a preset dropdown and a Save button at the top,
 * and a live preview allows adjusting parameters before applying the result.
 * You can save current settings as a session preset, and view the code snippet
 * for embedding. Preview objects are created on a temporary layer and applied
 * to the original object only when confirmed.
 *
 * 更新日: 2026-03-22
 * Updated: 2026-03-22
 */

var SCRIPT_VERSION = "v1.4.0";

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
    presetCustomPrefix: { ja: "カスタム", en: "Custom" },
    btnSave: { ja: "保存", en: "Save" },
    alertPresetName: { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
    alertPresetSaved: { ja: "プリセットを保存しました。\n\nコード組み込み用:\n", en: "Preset saved.\n\nCode snippet for embedding:\n" },
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
    panelPerforationCenter: { ja: "左右分割", en: "Center Divider" },
    panelDash: { ja: "分割線", en: "Divider Line" },
    panelEdge: { ja: "エッジ", en: "Edge" },
    panelCorner: { ja: "コーナー", en: "Corner" },
    chkEnable: { ja: "有効", en: "Enable" },
    lblLineWidth: { ja: "線幅:", en: "Weight:" },
    lblGap: { ja: "間隔:", en: "Gap:" },
    lblLength: { ja: "長さ:", en: "Inset Length:" },
    rdNone: { ja: "なし", en: "None" },
    rdDot: { ja: "ドット", en: "Dot" },
    rdDash: { ja: "破線", en: "Dash" },
    rdCircle: { ja: "円", en: "Circle" },
    rdTriangle: { ja: "三角", en: "Triangle" },
    rdWRound: { ja: "ダブル角丸", en: "Double Rounded" },
    rdRound: { ja: "角丸", en: "Rounded" },
    rdInverse: { ja: "逆角丸", en: "Inverse Round" },
    rdChamfer: { ja: "面取り", en: "Chamfer" },
    panelHole: { ja: "スリット／ホール", en: "Slit / Hole" },
    lblSize: { ja: "サイズ:", en: "Size:" },
    chkShapeOnly: { ja: "エッジのみ", en: "Edges Only" },
    chkLinkCenter: { ja: "分割線に連動", en: "Link to Split Line" },
    chkLeft: { ja: "左", en: "Left" },
    chkRight: { ja: "右", en: "Right" },
    rdLR: { ja: "左右", en: "L/R" },
    rdTB: { ja: "上下", en: "T/B" },
    lblZZSize: { ja: "大きさ:", en: "Size:" },
    lblZZRepeat: { ja: "繰り返し:", en: "Repeat:" },
    chkPreview: { ja: "プレビュー", en: "Preview" },
    chkExpandAppearance: { ja: "アピアランスを分割", en: "Expand Appearance" },
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOK: { ja: "OK", en: "OK" },
    btnOutlineOn: { ja: "アウトライン表示", en: "Outline View" },
    btnOutlineOff: { ja: "プレビュー表示", en: "Preview View" }
};
function cloneSimpleValue(value) {
    if (value === null || typeof value !== 'object') return value;
    if (value instanceof Array) {
        var arr = [];
        for (var i = 0; i < value.length; i++) arr.push(cloneSimpleValue(value[i]));
        return arr;
    }
    var out = {};
    for (var key in value) {
        if (value.hasOwnProperty(key)) out[key] = cloneSimpleValue(value[key]);
    }
    return out;
}


function isFiniteNumber(value) {
    return typeof value === 'number' && !isNaN(value) && isFinite(value);
}

function toCodeValue(value, indentLevel) {
    var indent = new Array(indentLevel + 1).join('    ');
    var nextIndent = new Array(indentLevel + 2).join('    ');
    var parts = [];
    var key;

    if (typeof value === 'string') {
        return '"' + value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r/g, '\\r').replace(/\n/g, '\\n') + '"';
    }
    if (typeof value === 'number') {
        return isFiniteNumber(value) ? String(value) : '0';
    }
    if (typeof value === 'boolean') {
        return value ? 'true' : 'false';
    }
    if (value === null) {
        return 'null';
    }
    if (value instanceof Array) {
        for (var i = 0; i < value.length; i++) {
            parts.push(toCodeValue(value[i], indentLevel + 1));
        }
        return '[' + parts.join(', ') + ']';
    }
    if (typeof value === 'object') {
        for (key in value) {
            if (value.hasOwnProperty(key)) {
                parts.push(nextIndent + key + ': ' + toCodeValue(value[key], indentLevel + 1));
            }
        }
        return '{\n' + parts.join(',\n') + '\n' + indent + '}';
    }
    return 'null';
}

function presetToCode(preset) {
    return ',\n' + toCodeValue(preset, 0);
}

var PRESETS = [
    {
        name: "0: クリア",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: false,
            offset: 30,
            mode: "dot",
            width: 3,
            gap: 6,
            length: -10
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "none",
            size: 5
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 10
        },
        edge: {
            mode: "wround",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    },

    {
        name: "1: W角丸＋分割線",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: true,
            offset: 30,
            mode: "dot",
            width: 3,
            gap: 6,
            length: -20
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "none",
            size: 5
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 10
        },
        edge: {
            mode: "wround",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    },
    {
        name: "2: 左右に2ホール",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: false,
            offset: 30,
            mode: "dot",
            width: 3,
            gap: 6,
            length: -6
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "none",
            size: 5
        },
        hole: {
            mode: "circle",
            left: true,
            right: true,
            size: 15
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    }, {
        name: "3: ミシン目",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: true,
            offset: 35,
            mode: "dot",
            width: 3,
            gap: 6,
            length: 0
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "round",
            size: 6
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 10
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    },
    {
        name: "4: 四隅に逆角丸",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: false,
            offset: 35,
            mode: "dot",
            width: 3,
            gap: 6,
            length: 0
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "inverse",
            size: 25
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 10
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    },

    {
        name: "5: 左右に三角スリット",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: false,
            offset: 35,
            mode: "dot",
            width: 3,
            gap: 6,
            length: 0
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 3
        },
        corner: {
            mode: "round",
            size: 3
        },
        hole: {
            mode: "triangle",
            left: true,
            right: true,
            size: 15
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    }
    ,
    {
        name: "7: 面取り＋分割線（破線）",
        lr: {
            enabled: false,
            linkCenter: false,
            width: 4,
            gap: 7,
            length: 0
        },
        center: {
            enabled: true,
            offset: 35,
            mode: "dash",
            width: 1,
            gap: 3,
            length: 0
        },
        zigzag: {
            mode: "none",
            size: 10,
            gap: 0,
            repeat: 7
        },
        corner: {
            mode: "chamfer",
            size: 13
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 15
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    },
    {
        name: "6: ミシン目＋ギザギザ上下",
        lr: {
            enabled: true,
            linkCenter: false,
            width: 4,
            gap: 7,
            length: 0
        },
        center: {
            enabled: false,
            offset: 35,
            mode: "dot",
            width: 3,
            gap: 6,
            length: 0
        },
        zigzag: {
            mode: "tb",
            size: 10,
            gap: 0,
            repeat: 7
        },
        corner: {
            mode: "none",
            size: 3
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 15
        },
        edge: {
            mode: "none",
            size: 5,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    }
    ,
    {
        name: "7: 分割線＋エッジ（円）＋ギザギザ",
        lr: {
            enabled: false,
            linkCenter: true,
            width: 3,
            gap: 6,
            length: 0
        },
        center: {
            enabled: true,
            offset: 30,
            mode: "dot",
            width: 3,
            gap: 6,
            length: -13
        },
        zigzag: {
            mode: "lr",
            size: 7,
            gap: 0,
            repeat: 7
        },
        corner: {
            mode: "none",
            size: 5
        },
        hole: {
            mode: "none",
            left: true,
            right: true,
            size: 10
        },
        edge: {
            mode: "circle",
            size: 10,
            shapeOnly: false
        },
        preview: {
            enabled: true,
            expandAppearance: false
        }
    }
];

function L(key) {
    var item = LABELS[key];
    if (!item) throw new Error("Missing label: " + key);
    return item[lang];
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
            var linked = ui.chkLR.value && ui.chkLinkCenter.value && ui.rdDotC.value;

            return {
                strokeWValueLR: linked ? parseFloat(ui.inputWidthC.text) : parseFloat(ui.inputWidthLR.text),
                gapValueLR: linked ? parseFloat(ui.inputGapC.text) : parseFloat(ui.inputGapLR.text),
                dashLenValueLR: (function (v) { v = linked ? parseFloat(ui.inputDashLenC.text) : parseFloat(ui.inputDashLenLR.text); return isNaN(v) ? v : Math.min(0, v); })(),
                lrLinked: linked,

                zzSizeValue: parseFloat(ui.inputWidthZZ.text),
                zzGapValue: parseFloat(ui.inputGapZZ.text) || 0,
                zzRepeatValue: Math.max(1, Math.round(parseFloat(ui.inputDashLenZZ.text) || 1)),

                strokeWValueC: parseFloat(ui.inputWidthC.text),
                gapValueC: parseFloat(ui.inputGapC.text),
                dashLenValueC: (function (v) { v = parseFloat(ui.inputDashLenC.text); return isNaN(v) ? v : Math.min(0, v); })(),
                offsetValue: parseFloat(ui.inputOffset.text),

                cornerSizeValue: parseFloat(ui.inputCornerSize.text),
                holeSizeValue: parseFloat(ui.inputHoleSize.text),
                edgeSizeValueC: parseFloat(ui.inputEdgeSizeC.text)
            };
        }

        function validateUIValues(values) {
            var numericKeys = [
                'strokeWValueLR', 'gapValueLR', 'dashLenValueLR',
                'zzSizeValue', 'zzGapValue', 'zzRepeatValue',
                'strokeWValueC', 'gapValueC', 'dashLenValueC', 'offsetValue',
                'cornerSizeValue', 'holeSizeValue', 'edgeSizeValueC'
            ];
            for (var i = 0; i < numericKeys.length; i++) {
                var key = numericKeys[i];
                if (isNaN(values[key])) return false;
            }
            return true;
        }

        // === Preset helpers (moved inside main) ===
        function getPresetDisplayName(preset) {
            if (!preset) return '';
            if (typeof preset.name === 'string') return preset.name;
            if (preset.name && typeof preset.name === 'object') {
                return preset.name[lang] || preset.name.ja || preset.name.en || '';
            }
            return '';
        }

        function getDropdownSelectionText(dropdown) {
            return (dropdown && dropdown.selection) ? dropdown.selection.text : '';
        }

        function setDropdownItemsFromPresets(dropdown, presets) {
            dropdown.removeAll();
            for (var i = 0; i < presets.length; i++) {
                dropdown.add('item', getPresetDisplayName(presets[i]));
            }
            if (dropdown.items.length > 0) dropdown.selection = 0;
        }

        function getSelectedPresetIndex() {
            return (ui.ddPreset && ui.ddPreset.selection) ? ui.ddPreset.selection.index : 0;
        }

        function buildPresetFromUI(name) {
            var presetName = name || getDropdownSelectionText(ui.ddPreset) || L('presetCustomPrefix');
            return {
                name: presetName,
                lr: {
                    enabled: ui.chkLR.value,
                    linkCenter: ui.chkLinkCenter.value,
                    width: parseFloat(ui.inputWidthLR.text),
                    gap: parseFloat(ui.inputGapLR.text),
                    length: parseFloat(ui.inputDashLenLR.text)
                },
                center: {
                    enabled: ui.chkCenter.value,
                    offset: parseFloat(ui.inputOffset.text),
                    mode: ui.rdDashC.value ? 'dash' : 'dot',
                    width: parseFloat(ui.inputWidthC.text),
                    gap: parseFloat(ui.inputGapC.text),
                    length: parseFloat(ui.inputDashLenC.text)
                },
                zigzag: {
                    mode: ui.rdZZLR.value ? 'lr' : (ui.rdZZTB.value ? 'tb' : 'none'),
                    size: parseFloat(ui.inputWidthZZ.text),
                    gap: parseFloat(ui.inputGapZZ.text),
                    repeat: parseFloat(ui.inputDashLenZZ.text)
                },
                corner: {
                    mode: ui.rdRound.value ? 'round' : (ui.rdInverse.value ? 'inverse' : (ui.rdChamfer.value ? 'chamfer' : 'none')),
                    size: parseFloat(ui.inputCornerSize.text)
                },
                hole: {
                    mode: ui.rdHoleCircle.value ? 'circle' : (ui.rdHoleTriangle.value ? 'triangle' : 'none'),
                    left: ui.chkHoleLeft.value,
                    right: ui.chkHoleRight.value,
                    size: parseFloat(ui.inputHoleSize.text)
                },
                edge: {
                    mode: ui.chkEdgeWRoundC.value ? 'wround' : (ui.rdEdgeCircleC.value ? 'circle' : (ui.rdEdgeTriangleC.value ? 'triangle' : 'none')),
                    size: parseFloat(ui.inputEdgeSizeC.text),
                    shapeOnly: ui.chkEdgeOnlyC.value
                },
                preview: {
                    enabled: ui.chkPreview.value,
                    expandAppearance: ui.chkExpandAppearance.value
                }
            };
        }

        function applyPresetToUI(preset) {
            if (!preset) return;

            ui.chkLR.value = !!preset.lr.enabled;
            ui.chkLinkCenter.value = !!preset.lr.linkCenter;
            ui.inputWidthLR.text = String(preset.lr.width);
            ui.inputGapLR.text = String(preset.lr.gap);
            ui.inputDashLenLR.text = String(preset.lr.length);

            ui.chkCenter.value = !!preset.center.enabled;
            ui.inputOffset.text = String(preset.center.offset);
            ui.rdDotC.value = preset.center.mode !== 'dash';
            ui.rdDashC.value = preset.center.mode === 'dash';
            ui.inputWidthC.text = String(preset.center.width);
            ui.inputGapC.text = String(preset.center.gap);
            ui.inputDashLenC.text = String(preset.center.length);

            ui.rdZZNone.value = preset.zigzag.mode === 'none';
            ui.rdZZLR.value = preset.zigzag.mode === 'lr';
            ui.rdZZTB.value = preset.zigzag.mode === 'tb';
            ui.inputWidthZZ.text = String(preset.zigzag.size);
            ui.inputGapZZ.text = String(preset.zigzag.gap);
            ui.inputDashLenZZ.text = String(preset.zigzag.repeat);

            ui.rdCornerNone.value = preset.corner.mode === 'none';
            ui.rdRound.value = preset.corner.mode === 'round';
            ui.rdInverse.value = preset.corner.mode === 'inverse';
            ui.rdChamfer.value = preset.corner.mode === 'chamfer';
            ui.inputCornerSize.text = String(preset.corner.size);

            ui.rdHoleNone.value = preset.hole.mode === 'none';
            ui.rdHoleCircle.value = preset.hole.mode === 'circle';
            ui.rdHoleTriangle.value = preset.hole.mode === 'triangle';
            ui.chkHoleLeft.value = !!preset.hole.left;
            ui.chkHoleRight.value = !!preset.hole.right;
            ui.inputHoleSize.text = String(preset.hole.size);

            ui.chkEdgeWRoundC.value = preset.edge.mode === 'wround';
            ui.rdEdgeNoneC.value = preset.edge.mode === 'none';
            ui.rdEdgeCircleC.value = preset.edge.mode === 'circle';
            ui.rdEdgeTriangleC.value = preset.edge.mode === 'triangle';
            ui.inputEdgeSizeC.text = String(preset.edge.size);
            ui.chkEdgeOnlyC.value = !!preset.edge.shapeOnly;

            ui.chkPreview.value = !!preset.preview.enabled;
            ui.chkExpandAppearance.value = !!preset.preview.expandAppearance;

            var offsetValue = parseFloat(ui.inputOffset.text);
            if (!isNaN(offsetValue)) {
                ui.sliderOffset.value = Math.max(-halfW_mm, Math.min(halfW_mm, offsetValue));
            }

            updateHoleDim();
            updateCornerDim();
            updateCenterDim();
            updateZZDim();
            updateLRDim();
            updateEdgeDimC();
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
            if (chkCenter.value && !(chkEdgeOnlyC.value && (!rdEdgeNoneC.value || chkEdgeWRoundC.value))) {
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
            if (chkCenter.value && !rdEdgeNoneC.value && !chkEdgeWRoundC.value) {
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

        function applyWRoundEdge(doc, rect, items, x, y, w, h, lineX, edgeSizeValueC, cornerSizeValue) {
            if (!chkCenter.value || !chkEdgeWRoundC.value) return false;
            if (isNaN(edgeSizeValueC) || edgeSizeValueC <= 0) return false;

            var edgeSizePtC = toPt(edgeSizeValueC, rulerPtFactor);

            // 分割線の位置で2つの長方形に分ける
            var leftW = lineX - x;
            var rightW = x + w - lineX;

            var leftRect = doc.pathItems.rectangle(y, x, leftW, h);
            leftRect.filled = true;
            leftRect.stroked = false;
            try { leftRect.fillColor = rect.fillColor; } catch (e) { }

            var rightRect = doc.pathItems.rectangle(y, lineX, rightW, h);
            rightRect.filled = true;
            rightRect.stroked = false;
            try { rightRect.fillColor = rect.fillColor; } catch (e) { }

            // それぞれにエッジサイズで角丸を適用
            var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + edgeSizePtC + ' "/></LiveEffect>';
            leftRect.applyEffect(xml);
            rightRect.applyEffect(xml);

            // コーナーパネルの角丸も追加適用
            if (rdRound.value && !isNaN(cornerSizeValue) && cornerSizeValue > 0) {
                var cornerSizePt = toPt(cornerSizeValue, rulerPtFactor);
                var cornerXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSizePt + ' "/></LiveEffect>';
                leftRect.applyEffect(cornerXml);
                rightRect.applyEffect(cornerXml);
            }

            // 元の長方形を置き換え
            rect.remove();
            var wGroup = groupItems(doc, [leftRect, rightRect]);
            selectItems(doc, [wGroup]);
            app.executeMenuCommand('Live Pathfinder Add');
            items[0] = getSingleSelection(doc);

            return true;
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

            var uiValues = collectUIValues();
            if (!validateUIValues(uiValues)) {
                throw new Error(L('alertEnterValidNumbers'));
            }

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

            // ミシン目（左右）— 連動時は分割線と同じ単位系を使用
            var lrStrokeFactor = uiValues.lrLinked ? strokePtFactor : rulerPtFactor;
            var strokeWLR = toPt(uiValues.strokeWValueLR, lrStrokeFactor);
            var gapLR = toPt(uiValues.gapValueLR, rulerPtFactor);
            var dashLenLR = toPt(Math.abs(uiValues.dashLenValueLR), rulerPtFactor);

            // ギザギザ
            var zzSizePt = toPt(uiValues.zzSizeValue, rulerPtFactor);
            var zzGapPt = toPt(uiValues.zzGapValue, rulerPtFactor);
            var zzRepeat = uiValues.zzRepeatValue;

            // ミシン目（中央）
            var strokeWC = toPt(uiValues.strokeWValueC, strokePtFactor);
            var gapC = toPt(uiValues.gapValueC, rulerPtFactor);
            var dashLenC = toPt(Math.abs(uiValues.dashLenValueC), rulerPtFactor);
            var offsetPt = toPt(uiValues.offsetValue, rulerPtFactor);

            // W角丸以外の場合のみコーナー角丸を先に適用
            var useWRound = chkCenter.value && chkEdgeWRoundC.value;
            if (!useWRound) {
                applyRoundCorners(targets, uiValues.cornerSizeValue);
            }

            for (var i = 0; i < targets.length; i++) {
                var rect = targets[i];
                var x = rect.left;
                var y = rect.top;
                var w = rect.width;
                var h = rect.height;
                var lineX = x + w / 2 + offsetPt;
                var items = [rect];

                // W角丸を適用（元の長方形を2分割して角丸）
                if (useWRound) {
                    applyWRoundEdge(doc, rect, items, x, y, w, h, lineX, uiValues.edgeSizeValueC, uiValues.cornerSizeValue);
                }

                // ミシン目を作成（中央）
                addCenterPerforation(doc, items, x, y, w, h, lineX, dashLenC, strokeWC, gapC);

                // ミシン目を作成（左右）
                addLRPerforation(doc, items, x, y, w, h, dashLenLR, strokeWLR, gapLR);

                // ギザギザを作成
                addZigzag(doc, items, x, y, w, h, zzSizePt, zzGapPt, zzRepeat);

                // エッジを作成（中央）
                addCenterEdges(doc, items, lineX, y, h, uiValues.edgeSizeValueC);

                // ホールを作成
                addHoles(doc, items, x, y, w, h, uiValues.holeSizeValue);

                // コーナーの逆角丸を作成
                addInverseCorners(doc, items, x, y, w, h, uiValues.cornerSizeValue);

                // コーナーの面取りを作成
                addChamferCorners(doc, items, x, y, w, h, uiValues.cornerSizeValue);

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
                // For length inputs (negative-only), cap at 0 as maximum
                if (targetInput === inputDashLenLR || targetInput === inputDashLenC) {
                    if (value > 0) value = 0;
                }

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

        var grpPresetTop = dlg.add('group');
        grpPresetTop.orientation = 'row';
        grpPresetTop.alignChildren = ['fill', 'center'];
        grpPresetTop.margins = [20, 0, 20, 0];
        grpPresetTop.alignment = ['fill', 'top'];
        grpPresetTop.spacing = 12;

        var ddPreset = grpPresetTop.add('dropdownlist', undefined, []);
        ddPreset.alignment = ['fill', 'center'];

        var btnSavePreset = grpPresetTop.add('button', undefined, L('btnSave'));
        btnSavePreset.preferredSize.width = 70;
        btnSavePreset.preferredSize.height = 24;
        btnSavePreset.alignment = ['right', 'center'];

        function getLabelWidth(base) {
            return (lang === 'ja') ? base : Math.round(base * 1.3);
        }

        var labelW = getLabelWidth(70);
        var labelW2 = getLabelWidth(78);

        // ===== 1行目: コーナー／ギザギザ =====
        var grpRow1 = dlg.add('group');
        grpRow1.orientation = 'row';
        grpRow1.alignChildren = ['fill', 'top'];

        var grpCol1 = grpRow1.add('group');
        grpCol1.orientation = 'column';
        grpCol1.alignChildren = ['fill', 'top'];

        var pnlCorner = grpCol1.add('panel', undefined, L('panelCorner'));
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
        grpCornerSize.add('statictext', undefined, rulerUnitLabel);

        // ===== ギザギザパネル =====
        var pnlZigzag = grpCol1.add('panel', undefined, L('panelZigzag'));
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
        grpWidthZZ.add('statictext', undefined, strokeUnitLabel);

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
        grpGapZZ.add('statictext', undefined, strokeUnitLabel);

        // ===== 左右パネル =====
        var pnlLRWrap = grpRow1.add('panel', undefined, L('panelLR'));
        pnlLRWrap.orientation = 'column';
        pnlLRWrap.alignChildren = ['fill', 'top'];
        pnlLRWrap.margins = [15, 20, 15, 15];

        // ===== ミシン目パネル =====
        var pnlLR = pnlLRWrap.add('panel', undefined, L('panelPerforationLR'));
        pnlLR.orientation = 'column';
        pnlLR.alignChildren = ['fill', 'top'];
        pnlLR.margins = [15, 20, 15, 10];

        var grpEnableLR = pnlLR.add('group');
        grpEnableLR.margins = [0, 0, 0, 0];
        var chkLR = grpEnableLR.add('checkbox', undefined, L('chkEnable'));
        var chkLinkCenter = grpEnableLR.add('checkbox', undefined, L('chkLinkCenter'));
        chkLinkCenter.value = true;

        var grpWidthLR = pnlLR.add('group');
        grpWidthLR.add('statictext', undefined, L('lblLineWidth'));
        var inputWidthLR = grpWidthLR.add('edittext', undefined, '3');
        inputWidthLR.characters = 3;
        changeValueByArrowKey(inputWidthLR, false, inputWidthLR);
        grpWidthLR.add('statictext', undefined, rulerUnitLabel);

        var grpGapLR = pnlLR.add('group');
        grpGapLR.add('statictext', undefined, L('lblGap'));
        var inputGapLR = grpGapLR.add('edittext', undefined, '6');
        inputGapLR.characters = 3;
        changeValueByArrowKey(inputGapLR, false, inputGapLR);
        grpGapLR.add('statictext', undefined, rulerUnitLabel);

        var grpDashLenLR = pnlLR.add('group');
        grpDashLenLR.add('statictext', undefined, L('lblLength'));
        var inputDashLenLR = grpDashLenLR.add('edittext', [0, 0, 40, 18], '0');
        changeValueByArrowKey(inputDashLenLR, true, inputDashLenLR);
        grpDashLenLR.add('statictext', undefined, rulerUnitLabel);

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
        grpHoleSize.add('statictext', undefined, rulerUnitLabel);

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
        chkCenter.value = true;
        var inputOffset = grpCenterTop.add('edittext', undefined, '0');
        inputOffset.characters = 5;
        changeValueByArrowKey(inputOffset, true, inputOffset);
        var lblOffsetMM = grpCenterTop.add('statictext', undefined, rulerUnitLabel);

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
        sliderOffset.preferredSize.width = 200;

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
        grpWidthC.add('statictext', undefined, strokeUnitLabel);

        var grpGapC = pnlDashC.add('group');
        grpGapC.add('statictext', undefined, L('lblGap'));
        var inputGapC = grpGapC.add('edittext', undefined, '6');
        inputGapC.characters = 3;
        changeValueByArrowKey(inputGapC, false, inputGapC);
        grpGapC.add('statictext', undefined, rulerUnitLabel);

        var grpDashLenC = pnlDashC.add('group');
        grpDashLenC.add('statictext', undefined, L('lblLength'));
        var inputDashLenC = grpDashLenC.add('edittext', [0, 0, 40, 18], '0');
        changeValueByArrowKey(inputDashLenC, true, inputDashLenC);
        grpDashLenC.add('statictext', undefined, rulerUnitLabel);

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

        var chkEdgeWRoundC = pnlEdgeC.add('checkbox', undefined, L('rdWRound'));

        var grpEdgeSizeC = pnlEdgeC.add('group');
        grpEdgeSizeC.add('statictext', undefined, L('lblSize'));
        var inputEdgeSizeC = grpEdgeSizeC.add('edittext', undefined, '10');
        inputEdgeSizeC.characters = 3;
        changeValueByArrowKey(inputEdgeSizeC, false, inputEdgeSizeC);
        grpEdgeSizeC.add('statictext', undefined, strokeUnitLabel);

        var chkEdgeOnlyC = pnlEdgeC.add('checkbox', undefined, L('chkShapeOnly'));

        // keep expand appearance checkbox separately (above or reuse if needed)
        var grpPreviewRow = dlg.add('group');
        grpPreviewRow.orientation = 'row';
        grpPreviewRow.alignment = ['center', 'top'];
        grpPreviewRow.alignChildren = ['left', 'center'];

        var chkPreview = grpPreviewRow.add('checkbox', undefined, L('chkPreview'));
        chkPreview.value = true;

        var chkExpandAppearance = grpPreviewRow.add('checkbox', undefined, L('chkExpandAppearance'));
        chkExpandAppearance.value = false;

        // Bottom button area (moved below preview row)
        var grpBottom = dlg.add('group');
        grpBottom.orientation = 'row';
        grpBottom.alignChildren = ['fill', 'center'];

        // Left column
        var grpLeft = grpBottom.add('group');
        grpLeft.orientation = 'row';
        grpLeft.alignChildren = ['left', 'center'];

        var isOutlineMode = false;
        var btnPreview = grpLeft.add('button', undefined, L('btnOutlineOn'));
        btnPreview.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isOutlineMode = !isOutlineMode;
                btnPreview.text = isOutlineMode ? L('btnOutlineOff') : L('btnOutlineOn');
            } catch (e) { }
        };

        // Right column
        var grpRight = grpBottom.add('group');
        grpRight.orientation = 'row';
        grpRight.alignment = ['right', 'center'];
        grpRight.alignChildren = ['right', 'center'];

        grpRight.add('button', undefined, L('btnCancel'), { name: 'cancel' });
        grpRight.add('button', undefined, L('btnOK'), { name: 'ok' });

        // UI参照の集約（段階的移行用）
        var ui = {
            // LR
            chkLR: chkLR,
            chkLinkCenter: chkLinkCenter,
            inputWidthLR: inputWidthLR,
            inputGapLR: inputGapLR,
            inputDashLenLR: inputDashLenLR,

            // Center
            chkCenter: chkCenter,
            inputOffset: inputOffset,
            sliderOffset: sliderOffset,
            rdDotC: rdDotC,
            rdDashC: rdDashC,
            inputWidthC: inputWidthC,
            inputGapC: inputGapC,
            inputDashLenC: inputDashLenC,

            // Zigzag
            rdZZNone: rdZZNone,
            rdZZLR: rdZZLR,
            rdZZTB: rdZZTB,
            inputWidthZZ: inputWidthZZ,
            inputGapZZ: inputGapZZ,
            inputDashLenZZ: inputDashLenZZ,

            // Corner
            rdCornerNone: rdCornerNone,
            rdRound: rdRound,
            rdInverse: rdInverse,
            rdChamfer: rdChamfer,
            inputCornerSize: inputCornerSize,

            // Hole
            rdHoleNone: rdHoleNone,
            rdHoleCircle: rdHoleCircle,
            rdHoleTriangle: rdHoleTriangle,
            chkHoleLeft: chkHoleLeft,
            chkHoleRight: chkHoleRight,
            inputHoleSize: inputHoleSize,

            // Edge
            rdEdgeNoneC: rdEdgeNoneC,
            rdEdgeCircleC: rdEdgeCircleC,
            rdEdgeTriangleC: rdEdgeTriangleC,
            chkEdgeWRoundC: chkEdgeWRoundC,
            inputEdgeSizeC: inputEdgeSizeC,
            chkEdgeOnlyC: chkEdgeOnlyC,

            // Preset
            ddPreset: ddPreset,
            btnSavePreset: btnSavePreset,

            // Preview
            chkPreview: chkPreview,
            chkExpandAppearance: chkExpandAppearance,

            // UI制御
            grpWidthLR: grpWidthLR,
            grpGapLR: grpGapLR,
            grpDashLenLR: grpDashLenLR,
            grpWidthZZ: grpWidthZZ,
            grpGapZZ: grpGapZZ,
            grpDashLenZZ: grpDashLenZZ,
            grpHoleSide: grpHoleSide,
            grpHoleSize: grpHoleSize,
            grpCornerSize: grpCornerSize,
            grpDashEdgeC: grpDashEdgeC,
            grpEdgeRadioC: grpEdgeRadioC,
            grpEdgeSizeC: grpEdgeSizeC,
            pnlDashC: pnlDashC,
            pnlCorner: pnlCorner,
            lblOffsetMM: lblOffsetMM
        };

        setDropdownItemsFromPresets(ddPreset, PRESETS);
        if (PRESETS.length > 0) {
            applyPresetToUI(cloneSimpleValue(PRESETS[0]));
        }

        /* プレビュー更新 / Update preview */
        function updatePreview() {
            updateZZDim();
            removePreviewLayer(chkPreview.value);
            if (!chkPreview.value) return;
            applyEffect(true);
        }

        /* イベント / Events */
        chkPreview.onClick = function () { updatePreview(); };

        ddPreset.onChange = function () {
            var index = getSelectedPresetIndex();
            if (index < 0 || index >= PRESETS.length) return;
            applyPresetToUI(cloneSimpleValue(PRESETS[index]));
            updatePreview();
        };

        btnSavePreset.onClick = function () {
            var baseName = getDropdownSelectionText(ui.ddPreset) || L('presetCustomPrefix');
            var newName = prompt(L('alertPresetName'), baseName);
            if (newName === null) return;
            newName = String(newName).replace(/^\s+|\s+$/g, '');
            if (!newName) return;

            if (!validateUIValues(collectUIValues())) {
                alert(L('alertEnterValidNumbers'));
                return;
            }

            var preset = buildPresetFromUI(newName);
            PRESETS.push(cloneSimpleValue(preset));
            setDropdownItemsFromPresets(ui.ddPreset, PRESETS);
            ui.ddPreset.selection = ui.ddPreset.items.length - 1;

            alert(L('alertPresetSaved') + presetToCode(preset));
        };

        // ミシン目（左右）イベント
        chkLR.onClick = function () {
            if (chkLR.value && rdZZLR.value) {
                rdZZNone.value = true;
                rdZZLR.value = false;
                updateZZDim();
            }
            updateLRDim();
            updatePreview();
        };
        inputWidthLR.onChanging = updatePreview;
        inputGapLR.onChanging = updatePreview;
        inputDashLenLR.onChanging = function () {
            var v = parseFloat(inputDashLenLR.text);
            if (!isNaN(v) && v > 0) {
                inputDashLenLR.text = '0';
            }
            updatePreview();
        };
        function syncLRFromCenter() {
            if (!chkLinkCenter.value || !rdDotC.value) return;
            inputWidthLR.text = inputWidthC.text;
            inputGapLR.text = inputGapC.text;
            inputDashLenLR.text = inputDashLenC.text;
        }
        function updateLRDim() {
            var on = chkLR.value;
            var linked = on && chkLinkCenter.value && rdDotC.value;
            chkLinkCenter.enabled = on;
            grpWidthLR.enabled = on && !linked;
            grpGapLR.enabled = on && !linked;
            grpDashLenLR.enabled = on && !linked;
            if (linked) syncLRFromCenter();
        }
        chkLinkCenter.onClick = function () {
            updateLRDim();
            updatePreview();
        };
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
        rdDotC.onClick = function () { updateLRDim(); updatePreview(); };
        rdDashC.onClick = function () { updateLRDim(); updatePreview(); };
        inputWidthC.onChanging = function () { syncLRFromCenter(); updatePreview(); };
        inputGapC.onChanging = function () { syncLRFromCenter(); updatePreview(); };
        inputDashLenC.onChanging = function () {
            var v = parseFloat(inputDashLenC.text);
            if (!isNaN(v) && v > 0) {
                inputDashLenC.text = '0';
            }
            syncLRFromCenter();
            updatePreview();
        };
        function updateCenterDim() {
            var on = ui.chkCenter.value;
            ui.inputOffset.enabled = on;
            ui.lblOffsetMM.enabled = on;
            ui.sliderOffset.enabled = on;
            ui.grpDashEdgeC.enabled = on;
        }
        function syncDashLenFromEdgeSizeC() {
            if (!chkCenter.value || (rdEdgeNoneC.value && !chkEdgeWRoundC.value)) return;

            var edgeSizeValue = parseFloat(inputEdgeSizeC.text);
            if (isNaN(edgeSizeValue)) return;

            var dashLenValue = edgeSizeValue * 1.2;
            if (rdEdgeCircleC.value) {
                dashLenValue = edgeSizeValue * 1.0;
            } else if (chkEdgeWRoundC.value) {
                dashLenValue = edgeSizeValue * 2.0;
            }

            dashLenValue = Math.round(dashLenValue * 1000) / 1000;
            inputDashLenC.text = String(-dashLenValue);
        }

        function updateEdgeDimC() {
            var onEdge = ui.chkCenter.value && !ui.rdEdgeNoneC.value;
            var onWRound = ui.chkCenter.value && ui.chkEdgeWRoundC.value;
            var on = onEdge || onWRound;
            ui.grpEdgeRadioC.enabled = !onWRound;
            ui.grpEdgeSizeC.enabled = on;
            ui.chkEdgeOnlyC.enabled = on;
            ui.pnlDashC.enabled = ui.chkCenter.value && !(on && ui.chkEdgeOnlyC.value);
            if (on) syncDashLenFromEdgeSizeC();
        }
        updateCenterDim();
        updateEdgeDimC();
        rdEdgeNoneC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        rdEdgeCircleC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        rdEdgeTriangleC.onClick = function () { updateEdgeDimC(); updatePreview(); };
        chkEdgeWRoundC.onClick = function () {
            if (chkEdgeWRoundC.value) {
                // コーナーを「なし」に
                rdCornerNone.value = true;
                rdRound.value = false;
                rdInverse.value = false;
                rdChamfer.value = false;
                updateCornerDim();
                // エッジを「なし」に
                rdEdgeNoneC.value = true;
                rdEdgeCircleC.value = false;
                rdEdgeTriangleC.value = false;
                inputEdgeSizeC.text = '5';
            }
            else {
                // ダブル角丸をOFFにしたら、コーナーパネルのディムを解除し、
                // ON時に強制で「なし」へ変更していた場合は通常の角丸に戻す
                if (rdCornerNone.value) {
                    rdCornerNone.value = false;
                    rdRound.value = true;
                    rdInverse.value = false;
                    rdChamfer.value = false;
                }
                pnlCorner.enabled = true;
                grpCornerSize.enabled = !rdCornerNone.value;
            }
            syncDashLenFromEdgeSizeC();
            updateEdgeDimC();
            updatePreview();
        };
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
            var wRoundOn = ui.chkEdgeWRoundC.value;
            ui.pnlCorner.enabled = !wRoundOn;
            ui.grpCornerSize.enabled = !wRoundOn && !ui.rdCornerNone.value;
        }
        updateCornerDim();
        rdCornerNone.onClick = function () { updateCornerDim(); updatePreview(); };
        rdRound.onClick = function () { updateCornerDim(); updatePreview(); };
        rdInverse.onClick = function () { updateCornerDim(); updatePreview(); };
        rdChamfer.onClick = function () { updateCornerDim(); updatePreview(); };
        inputCornerSize.onChanging = updatePreview;

        syncDashLenFromEdgeSizeC();
        updatePreview();
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

main()