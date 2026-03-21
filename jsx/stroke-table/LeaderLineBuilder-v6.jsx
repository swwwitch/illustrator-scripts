#target illustrator
#targetengine "AngledLeaderLineMakerEngine"

#include "ColorPicker.jsx"

/*
 * 角度指定の引き出し線メーカー.jsx
 * v1.5.1
 * 更新日: 20260310
 *
 * 選択したパスまたはグループから、指定角度の引き出し線を生成するスクリプトです。
 * リアルタイムプレビュー、線端の丸、白フチ、線色、線幅、グループ化に対応しています。
 * 1.5.1では、ダイアログ構築・イベント登録・設定復元・設定保存を4分割し、責務を整理しました。
 */

var SCRIPT_VERSION = "v1.5.1";

// セッション中の設定を記憶
if (typeof $.global._leaderLineSettings === "undefined") {
    $.global._leaderLineSettings = {
        angle: "45",
        radioAngle: 45,
        applyScope: "all",
        hasUserSetApplyScope: false,
        diagDir: "upperLeft",
        hDir: "left",
        vDir: "up",
        capType: "circle",
        capStyle: "fill",
        capSize: "3",
        strokeCapType: "none",
        groupEnabled: true,
        whiteEdge: false,
        edgeColor: "white",
        edgeColorHex: "#ffcc00",
        dialogBounds: null, // ダイアログ位置 [x, y] のみをセッション中に保存
        lineColor: "black",
        lineColorHex: "#ffcc00",
        lineWidth: "1"
    };
}

// 現在のロケールを取得 / Get current locale
function getCurrentLang() {
    var locale = $.locale || '';
    if (locale.indexOf("ja") === 0) return "ja";
    return "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "角度指定の引き出し線メーカー",
        en: "Angled Leader Line Maker"
    },
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNoSelection: {
        ja: "パスまたはグループを選択して実行してください。",
        en: "Please select a path or group and run the script."
    },
    alertNoValidTargets: {
        ja: "パス（2点以上）またはグループが選択されていません。",
        en: "No valid path (2 or more points) or group is selected."
    },
    alertInvalidAngle: {
        ja: "角度は0より大きく90未満の値を入力してください。",
        en: "Enter an angle greater than 0 and less than 90."
    },
    panelAngle: {
        ja: "角度",
        en: "Angle"
    },
    panelDirection: {
        ja: "斜線の方向",
        en: "Diagonal Direction"
    },
    panelApplyScope: {
        ja: "適用範囲",
        en: "Apply Scope"
    },
    scopeAll: {
        ja: "すべて",
        en: "All"
    },
    scopeExceptDirection: {
        ja: "斜線の方向以外",
        en: "Except Direction"
    },
    panelLineStyle: {
        ja: "線のスタイル",
        en: "Line Style"
    },
    panelLineEnd: {
        ja: "線端",
        en: "Line End"
    },
    capShapeLabel: {
        ja: "形状",
        en: "Shape"
    },
    capStrokeCapLabel: {
        ja: "線端",
        en: "Stroke Cap"
    },
    capStrokeCapNone: {
        ja: "なし",
        en: "None"
    },
    panelOptions: {
        ja: "フチ",
        en: "Edge"
    },
    dirUpperLeft: {
        ja: "左上",
        en: "Upper Left"
    },
    dirLowerLeft: {
        ja: "左下",
        en: "Lower Left"
    },
    dirUpperRight: {
        ja: "右上",
        en: "Upper Right"
    },
    dirLowerRight: {
        ja: "右下",
        en: "Lower Right"
    },
    colorBlack: {
        ja: "黒",
        en: "Black"
    },
    colorWhite: {
        ja: "白",
        en: "White"
    },
    colorOther: {
        ja: "その他",
        en: "Other"
    },
    lineWidth: {
        ja: "線幅",
        en: "Line Width"
    },
    capNone: {
        ja: "なし",
        en: "None"
    },
    capRound: {
        ja: "丸型",
        en: "Round"
    },
    capCircle: {
        ja: "円",
        en: "Circle"
    },
    capArrow: {
        ja: "矢印",
        en: "Arrow"
    },
    capFill: {
        ja: "塗り",
        en: "Fill"
    },
    capStroke: {
        ja: "線",
        en: "Stroke"
    },
    capSize: {
        ja: "大きさ",
        en: "Size"
    },
    groupEnabled: {
        ja: "グループ化",
        en: "Group"
    },
    whiteEdge: {
        ja: "フチ",
        en: "White Edge"
    },
    zoom: {
        ja: "ズーム",
        en: "Zoom"
    },
    lightMode: {
        ja: "軽",
        en: "Light mode"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    }
};

/* ラベル取得 / Get localized label */
function L(key) {
    var item = LABELS[key];
    if (!item) return key;
    return item[lang] || item.en || key;
}

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

/* 単位コードと設定キーから適切な単位ラベルを返す / Get unit label from unit code and preference key */
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

/* 設定キーから現在の単位ラベルを取得 / Get current unit label from preference key */
function getCurrentUnitLabel(prefKey) {
    var unitCode = app.preferences.getIntegerPreference(prefKey);
    return getUnitLabel(unitCode, prefKey);
}

/* 単位コードからpt換算係数を取得 / Get pt conversion factor from unit code */
function getUnitToPtFactor(code, prefKey) {
    switch (code) {
        case 0: return 72;          // in
        case 1: return 72 / 25.4;   // mm
        case 2: return 1;           // pt
        case 3: return 12;          // pica
        case 4: return 72 / 2.54;   // cm
        case 5: return 0.25;        // Q/H
        case 6: return 1;           // px
        case 7: return 72;          // ft/in（数値入力はinch扱い）
        case 8: return 72 / 0.0254; // m
        case 9: return 72 * 36;     // yd
        case 10: return 72 * 12;    // ft
        default: return 1;
    }
}

/* 設定キーから現在のpt換算係数を取得 / Get current pt conversion factor from preference key */
function getCurrentUnitToPtFactor(prefKey) {
    var unitCode = app.preferences.getIntegerPreference(prefKey);
    return getUnitToPtFactor(unitCode, prefKey);
}

/* 単位値をptに変換 / Convert unit value to points */
function unitValueToPt(value, prefKey) {
    var n = parseFloat(value);
    if (isNaN(n)) return NaN;
    return n * getCurrentUnitToPtFactor(prefKey);
}

/* ptを単位値に変換 / Convert points to unit value */
function ptValueToUnit(valuePt, prefKey) {
    var n = parseFloat(valuePt);
    if (isNaN(n)) return NaN;
    return n / getCurrentUnitToPtFactor(prefKey);
}

function roundTo(value, digits) {
    var p = Math.pow(10, digits || 0);
    return Math.round(value * p) / p;
}

function formatNumber(value, digits) {
    if (isNaN(value)) return "";
    var s = String(roundTo(value, digits || 0));
    s = s.replace(/(\.\d*?)0+$/, "$1");
    // digits>=1 のとき最低1桁の小数を保持（1 → 1.0）
    if ((digits || 0) >= 1 && s.indexOf(".") === -1) s += ".0";
    return s;
}

function formatUnitValue(valuePt, prefKey) {
    return formatNumber(ptValueToUnit(valuePt, prefKey), 2);
}

// =========================================
// TMK Zoom Module (collision-safe + Light mode)
// - Light mode: apply zoom only on slider release
// =========================================
function __TMKZoom_captureViewState(doc) {
    var st = { view: null, zoom: null, center: null };
    try {
        st.view = doc.activeView;
        st.zoom = st.view.zoom;
        st.center = st.view.centerPoint;
    } catch (_) { }
    return st;
}

function __TMKZoom_restoreViewState(doc, state) {
    if (!state) return;
    try {
        var v = state.view || doc.activeView;
        if (v && state.zoom != null) v.zoom = state.zoom;
        if (v && state.center != null) v.centerPoint = state.center;
    } catch (_) { }
}

function __TMKZoom_addControls(parent, doc, labelText, initialState, options) {
    options = options || {};
    var minZoom = (typeof options.min === "number") ? options.min : 0.1;
    var maxZoom = (typeof options.max === "number") ? options.max : 16;
    var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : 240;
    var doRedraw = (options.redraw !== false);

    var showLightMode = (options.lightMode !== false);
    var lightModeLabel = options.lightModeLabel || "Light mode";
    var lightModeDefault = (options.lightModeDefault === true);

    var g = parent.add("group");
    g.orientation = "row";
    g.alignChildren = ["center", "center"];
    g.alignment = "center";
    try { if (options.margins) g.margins = options.margins; } catch (_) { }

    var stLabel = g.add("statictext", undefined, String(labelText || "Zoom"));

    var initZoom = 1;
    try {
        if (initialState && initialState.zoom != null) initZoom = Number(initialState.zoom);
        else initZoom = Number(doc.activeView.zoom);
    } catch (_) { }
    if (!initZoom || isNaN(initZoom)) initZoom = 1;

    var sld = g.add("slider", undefined, initZoom, minZoom, maxZoom);
    try { sld.preferredSize.width = sliderWidth; } catch (_) { }

    var chkLight = null;
    if (showLightMode) {
        chkLight = g.add("checkbox", undefined, String(lightModeLabel));
        chkLight.value = lightModeDefault;
    }

    function isLightMode() {
        return !!(chkLight && chkLight.value);
    }

    function applyZoom(z) {
        try {
            var v = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!v) return;
            v.zoom = z;
            if (doRedraw) { try { app.redraw(); } catch (_) { } }
        } catch (_) { }
    }

    function syncFromView() {
        try {
            var v = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!v) return;
            sld.value = v.zoom;
        } catch (_) { }
    }

    sld.onChanging = function () {
        if (isLightMode()) return;
        applyZoom(Number(sld.value));
    };

    sld.onChange = function () {
        applyZoom(Number(sld.value));
    };

    if (chkLight) {
        chkLight.onClick = function () {
            try { applyZoom(Number(sld.value)); } catch (_) { }
        };
    }

    return {
        group: g,
        label: stLabel,
        slider: sld,
        lightModeCheckbox: chkLight,
        applyZoom: applyZoom,
        syncFromView: syncFromView,
        restoreInitial: function () { __TMKZoom_restoreViewState(doc, initialState); }
    };
}

function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(L("alertNoDocument"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // 選択オブジェクトのチェック
    if (sel.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    // 対象を収集（PathItem：2点以上、GroupItem：通常は全体参照、leader_line 再適用時は A線優先）
    var targets = [];
    for (var i = 0; i < sel.length; i++) {
        var t = sel[i].typename;
        if (t === "PathItem" && sel[i].pathPoints.length >= 2) {
            targets.push(sel[i]);
        } else if (t === "GroupItem") {
            targets.push(sel[i]);
        }
    }

    if (targets.length === 0) {
        alert(L("alertNoValidTargets"));
        return;
    }

    // GroupItem の参照用 PathItem を取得する関数
    // 通常は closed path を優先して返す（leader_line 再適用時の A線優先ロジックは getLeaderLineBasePath() 側で処理）
    function getGroupReferencePathItems(groupItem) {
        var closedItems = [];
        var allPathItems = [];
        for (var k = 0; k < groupItem.pathItems.length; k++) {
            var pi = groupItem.pathItems[k];
            allPathItems.push(pi);
            if (pi.closed) closedItems.push(pi);
        }
        return closedItems.length ? closedItems : allPathItems;
    }

    // PathItem のアンカー座標から外接矩形を取得する関数
    function getPathAnchorBounds(pathItem) {
        var pts = pathItem.pathPoints;
        if (!pts || pts.length === 0) {
            return {
                x_left: 0,
                x_right: 0,
                y_top: 0,
                y_bottom: 0
            };
        }

        var xMin = pts[0].anchor[0], xMax = pts[0].anchor[0];
        var yMin = pts[0].anchor[1], yMax = pts[0].anchor[1];
        for (var i = 1; i < pts.length; i++) {
            var ax = pts[i].anchor[0];
            var ay = pts[i].anchor[1];
            if (ax < xMin) xMin = ax;
            if (ax > xMax) xMax = ax;
            if (ay < yMin) yMin = ay;
            if (ay > yMax) yMax = ay;
        }

        return {
            x_left: xMin,
            x_right: xMax,
            y_top: yMax,
            y_bottom: yMin
        };
    }

    // GroupItem の bounds を取得する関数 / 通常はグループ全体、leader_line 再適用時は保存済み bounds を優先し、なければ旧構造から復元する
    function getGroupBounds(groupItem) {
        if (isLeaderLineGroup(groupItem)) {
            var storedBounds = getLeaderLineStoredBounds(groupItem);
            if (storedBounds) {
                return storedBounds;
            }

            var basePath = getLeaderLineBasePath(groupItem);
            if (basePath) {
                var mainCap = getLeaderLineMainCap(groupItem);
                return getPathAnchorBoundsWithLeaderCap(basePath, mainCap);
            }
        }

        var gb = groupItem.geometricBounds; // [left, top, right, bottom]
        return {
            x_left: gb[0],
            x_right: gb[2],
            y_top: gb[1],
            y_bottom: gb[3]
        };
    }
    function getLeaderLineMainCap(groupItem) {
        for (var i = 0; i < groupItem.pathItems.length; i++) {
            var pi = groupItem.pathItems[i];
            if (pi.note === "leader_line_main_cap") return pi;
        }
        return null;
    }

    function getCapCenter(capItem) {
        if (!capItem) return null;
        var capGB = capItem.geometricBounds; // [left, top, right, bottom]
        return [
            (capGB[0] + capGB[2]) / 2,
            (capGB[1] + capGB[3]) / 2
        ];
    }

    function getPathAnchorBoundsWithLeaderCap(pathItem, capItem) {
        var bounds = getPathAnchorBounds(pathItem);
        if (!capItem) return bounds;

        var pts = pathItem.pathPoints;
        if (!pts || pts.length < 2) return bounds;

        var p0 = pts[0].anchor;
        var p1 = pts[1].anchor;
        var p2 = pts[pts.length - 1].anchor;
        var capCenter = getCapCenter(capItem);
        if (!capCenter) return bounds;

        // 線端に丸がある場合、丸の中心（= 丸の半分位置）を期待する端として扱う
        var cx = capCenter[0];
        var cy = capCenter[1];

        var d0 = Math.abs(p0[0] - cx) + Math.abs(p0[1] - cy);
        var d2 = Math.abs(p2[0] - cx) + Math.abs(p2[1] - cy);

        var tip0 = [p0[0], p0[1]];
        var bend = [p1[0], p1[1]];
        var tip2 = [p2[0], p2[1]];

        if (d0 <= d2) {
            tip0 = [cx, cy];
        } else {
            tip2 = [cx, cy];
        }

        return {
            x_left: Math.min(tip0[0], bend[0], tip2[0]),
            x_right: Math.max(tip0[0], bend[0], tip2[0]),
            y_top: Math.max(tip0[1], bend[1], tip2[1]),
            y_bottom: Math.min(tip0[1], bend[1], tip2[1])
        };
    }

    // 線の属性を取得する関数（GroupItem は通常 closed path 優先、leader_line 再適用時は A線優先）
    function getStrokeInfo(item) {
        if (item.typename === "PathItem") {
            return {
                stroked: item.stroked,
                strokeWidth: item.stroked ? item.strokeWidth : 1,
                strokeColor: item.stroked ? item.strokeColor : undefined
            };
        }
        // GroupItemの場合、closed pathを優先して配下のPathItemを探す
        if (isLeaderLineGroup(item)) {
            var leaderBasePath = getLeaderLineBasePath(item);
            if (leaderBasePath) {
                return {
                    stroked: leaderBasePath.stroked,
                    strokeWidth: leaderBasePath.stroked ? leaderBasePath.strokeWidth : 1,
                    strokeColor: leaderBasePath.stroked ? leaderBasePath.strokeColor : undefined
                };
            }
        }

        // GroupItemの場合、closed pathを優先して配下のPathItemを探す
        var refItems = getGroupReferencePathItems(item);
        for (var k = 0; k < refItems.length; k++) {
            var pi = refItems[k];
            if (pi.stroked) {
                return {
                    stroked: true,
                    strokeWidth: pi.strokeWidth,
                    strokeColor: pi.strokeColor
                };
            }
        }
        return { stroked: false, strokeWidth: 1, strokeColor: undefined };
    }

    // 各ターゲットの座標情報を都度取得する関数
    function collectBoundsData(targetItems) {
        var data = [];
        for (var i = 0; i < targetItems.length; i++) {
            var item = targetItems[i];
            var x_left, x_right, y_bottom, y_top;

            if (item.typename === "PathItem") {
                // PathItem：アンカーポイントから外接矩形を算出
                var pathBounds = getPathAnchorBounds(item);
                x_left = pathBounds.x_left;
                x_right = pathBounds.x_right;
                y_bottom = pathBounds.y_bottom;
                y_top = pathBounds.y_top;
            } else {
                // GroupItem：通常はグループ全体、leader_line 再適用時は保存済み bounds を優先して使用
                var groupBounds = getGroupBounds(item);
                x_left = groupBounds.x_left;
                x_right = groupBounds.x_right;
                y_top = groupBounds.y_top;
                y_bottom = groupBounds.y_bottom;
            }

            var si = getStrokeInfo(item);
            data.push({
                target: item,
                x_left: x_left,
                x_right: x_right,
                y_bottom: y_bottom,
                y_top: y_top,
                stroked: si.stroked,
                strokeWidth: si.strokeWidth,
                strokeColor: si.strokeColor
            });
        }
        return data;
    }

    // プレビュー用アイテムの配列（親グループ単位で管理）
    var previewItems = [];

    // Illustratorアイテムを安全に削除するヘルパー
    function safeRemove(item) {
        if (!item) return;
        try {
            if (item.parent) item.remove();
        } catch (_) { }
    }

    function safeSelect(item) {
        if (!item) return;
        try {
            item.selected = true;
        } catch (_) { }
    }

    function safeSetNote(item, note) {
        if (!item) return;
        try {
            item.note = note;
        } catch (_) { }
    }

    // HEXカラーをRGBColorに変換する関数
    function hexToRGBColor(hex) {
        hex = hex.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var r = parseInt(hex.substring(0, 2), 16);
        var g = parseInt(hex.substring(2, 4), 16);
        var b = parseInt(hex.substring(4, 6), 16);
        if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
        var color = new RGBColor();
        color.red = r;
        color.green = g;
        color.blue = b;
        return color;
    }

    function isCMYKDocument() {
        return doc.documentColorSpace === DocumentColorSpace.CMYK;
    }

    function createBlackColor() {
        if (isCMYKDocument()) {
            var black = new CMYKColor();
            black.cyan = 0;
            black.magenta = 0;
            black.yellow = 0;
            black.black = 100;
            return black;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    function createWhiteColor() {
        if (isCMYKDocument()) {
            var white = new CMYKColor();
            white.cyan = 0;
            white.magenta = 0;
            white.yellow = 0;
            white.black = 0;
            return white;
        }
        var rgbWhite = new RGBColor();
        rgbWhite.red = 255;
        rgbWhite.green = 255;
        rgbWhite.blue = 255;
        return rgbWhite;
    }

    function rgbToCMYKColor(r, g, b) {
        var rr = Math.max(0, Math.min(255, r)) / 255;
        var gg = Math.max(0, Math.min(255, g)) / 255;
        var bb = Math.max(0, Math.min(255, b)) / 255;
        var k = 1 - Math.max(rr, gg, bb);
        var c = 0, m = 0, y = 0;

        if (k < 1) {
            c = (1 - rr - k) / (1 - k);
            m = (1 - gg - k) / (1 - k);
            y = (1 - bb - k) / (1 - k);
        }

        var color = new CMYKColor();
        color.cyan = Math.round(c * 100);
        color.magenta = Math.round(m * 100);
        color.yellow = Math.round(y * 100);
        color.black = Math.round(k * 100);
        return color;
    }

    function hexToDocumentColor(hex) {
        var rgb = hexToRGBColor(hex);
        if (!rgb) return null;
        if (isCMYKDocument()) {
            return rgbToCMYKColor(rgb.red, rgb.green, rgb.blue);
        }
        return rgb;
    }

    function updateSwatch() {
        if (!colorSwatch || !hexInput) return;
        var hex = String(hexInput.text || "").replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var r = parseInt(hex.substring(0, 2), 16) / 255;
        var g = parseInt(hex.substring(2, 4), 16) / 255;
        var b = parseInt(hex.substring(4, 6), 16) / 255;
        if (isNaN(r) || isNaN(g) || isNaN(b)) return;
        var gfx = colorSwatch.graphics;
        gfx.backgroundColor = gfx.newBrush(gfx.BrushType.SOLID_COLOR, [r, g, b]);
    }

    function updateEdgeSwatch() {
        if (!edgeSwatch || !edgeHexInput) return;
        var hex = String(edgeHexInput.text || "").replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var r = parseInt(hex.substring(0, 2), 16) / 255;
        var g = parseInt(hex.substring(2, 4), 16) / 255;
        var b = parseInt(hex.substring(4, 6), 16) / 255;
        if (isNaN(r) || isNaN(g) || isNaN(b)) return;
        var gfx = edgeSwatch.graphics;
        gfx.backgroundColor = gfx.newBrush(gfx.BrushType.SOLID_COLOR, [r, g, b]);
    }

    // フチのカラーを取得する関数
    function getEdgeColor() {
        if (edgeColorWhite.value) {
            return createWhiteColor();
        } else {
            var docColor = hexToDocumentColor(edgeHexInput.text);
            if (docColor) return docColor;
            return createWhiteColor();
        }
    }

    // 線のカラーを取得する関数
    function getLineColor(b) {
        if (colorBlack.value) {
            return createBlackColor();
        } else if (colorWhite.value) {
            return createWhiteColor();
        } else {
            // その他：HEXカラーコードから現在のドキュメント色空間に合わせて変換
            var docColor = hexToDocumentColor(hexInput.text);
            if (docColor) return docColor;
            // 無効な値の場合は元のオブジェクトの線色を使用
            if (b.strokeColor) return b.strokeColor;
            return createBlackColor();
        }
    }

    // 線幅を取得する関数
    function getLineWidth(b) {
        var w = unitValueToPt(lineWidthInput.text, "strokeUnits");
        if (!isNaN(w) && w > 0) return w;
        return b.strokeWidth;
    }

    // 引き出し線の座標を計算する関数
    function calcLeaderPoints(b, angleRad, hDir, vDir) {
        var height = b.y_top - b.y_bottom;
        var offset = height / Math.tan(angleRad);
        if (vDir === "down") {
            // 水平線が上、斜線が下に向かう
            if (hDir === "left") {
                var ix = b.x_left + offset;
                return [[b.x_left, b.y_bottom], [ix, b.y_top], [b.x_right, b.y_top]];
            } else {
                var ix = b.x_right - offset;
                return [[b.x_left, b.y_top], [ix, b.y_top], [b.x_right, b.y_bottom]];
            }
        } else {
            // 水平線が下、斜線が上に向かう
            if (hDir === "left") {
                var ix = b.x_left + offset;
                return [[b.x_left, b.y_top], [ix, b.y_bottom], [b.x_right, b.y_bottom]];
            } else {
                var ix = b.x_right - offset;
                return [[b.x_left, b.y_bottom], [ix, b.y_bottom], [b.x_right, b.y_top]];
            }
        }
    }

    // 斜線の先端座標を取得する関数
    function getTipPoint(points, hDir) {
        return (hDir === "left") ? points[0] : points[2];
    }

    // 斜線の先端を円の縁まで短縮する関数
    function shortenTip(points, hDir, r) {
        var tipIdx = (hDir === "left") ? 0 : 2;
        var bendIdx = 1;
        var tip = points[tipIdx];
        var bend = points[bendIdx];
        var dx = bend[0] - tip[0];
        var dy = bend[1] - tip[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len > 0) {
            points[tipIdx] = [tip[0] + dx / len * r, tip[1] + dy / len * r];
        }
    }

    // 先端に円を追加する関数
    function addTipCircle(tipPt, b, diameter, fillOnly) {
        var r = diameter / 2;
        var lineColor = getLineColor(b);
        var sw = getLineWidth(b);
        var circle = doc.pathItems.ellipse(
            tipPt[1] + r, tipPt[0] - r, diameter, diameter
        );
        if (fillOnly) {
            // ●（塗りのみ）
            circle.filled = true;
            circle.fillColor = lineColor;
            circle.stroked = false;
        } else {
            // ○（線のみ）
            circle.filled = false;
            circle.stroked = true;
            circle.strokeWidth = sw;
            circle.strokeColor = lineColor;
        }
        return circle;
    }

    // 先端の円にフチを追加する関数
    function addTipCircleWhiteEdge(tipPt, b, diameter, fillOnly) {
        var sw = getLineWidth(b);
        var whiteColor = getEdgeColor();
        if (fillOnly) {
            // ●のフチ：塗りを白で大きめ円
            var expand = sw * 2;
            var newDiam = diameter + expand;
            var newR = newDiam / 2;
            var whiteCircle = doc.pathItems.ellipse(
                tipPt[1] + newR, tipPt[0] - newR, newDiam, newDiam
            );
            whiteCircle.filled = true;
            whiteCircle.fillColor = whiteColor;
            whiteCircle.stroked = false;
        } else {
            // ○のフチ：線ではなく白い塗り円を背面に置く
            var expand = sw * 2;
            var newDiam = diameter + expand;
            var newR = newDiam / 2;
            var whiteCircle = doc.pathItems.ellipse(
                tipPt[1] + newR, tipPt[0] - newR, newDiam, newDiam
            );
            whiteCircle.filled = true;
            whiteCircle.fillColor = whiteColor;
            whiteCircle.stroked = false;
        }
        return whiteCircle;
    }

    // 先端に矢印を追加する関数
    function addTipArrow(tipPt, bendPt, b, arrowSize, fillOnly) {
        var lineColor = getLineColor(b);
        var sw = getLineWidth(b);
        // tipPt → bendPt 方向のベクトル
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) return null;
        var ux = dx / len;
        var uy = dy / len;
        // 矢印の2つの翼端を計算
        var halfW = arrowSize / 2;
        var backX = tipPt[0] + ux * arrowSize;
        var backY = tipPt[1] + uy * arrowSize;
        var wing1 = [backX + uy * halfW, backY - ux * halfW];
        var wing2 = [backX - uy * halfW, backY + ux * halfW];

        var arrow = doc.pathItems.add();
        arrow.setEntirePath([tipPt, wing1, wing2]);
        arrow.closed = true;
        if (fillOnly) {
            arrow.filled = true;
            arrow.fillColor = lineColor;
            arrow.stroked = false;
        } else {
            arrow.filled = false;
            arrow.stroked = true;
            arrow.strokeWidth = sw;
            arrow.strokeColor = lineColor;
        }
        return arrow;
    }

    // 先端の矢印にフチを追加する関数（本体と同じ重心を基準にスケールアップ）
    function addTipArrowWhiteEdge(tipPt, bendPt, b, arrowSize, fillOnly) {
        var sw = getLineWidth(b);
        var whiteColor = getEdgeColor();
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) return null;
        var ux = dx / len;
        var uy = dy / len;
        // 本体の矢印の3頂点を計算
        var halfW = arrowSize / 2;
        var backX = tipPt[0] + ux * arrowSize;
        var backY = tipPt[1] + uy * arrowSize;
        var p0 = tipPt;
        var p1 = [backX + uy * halfW, backY - ux * halfW];
        var p2 = [backX - uy * halfW, backY + ux * halfW];
        // 重心を求める
        var cx = (p0[0] + p1[0] + p2[0]) / 3;
        var cy = (p0[1] + p1[1] + p2[1]) / 3;
        // 重心を基準にスケールアップ
        var scale = (arrowSize + sw * 3) / arrowSize;
        var fp0 = [cx + (p0[0] - cx) * scale, cy + (p0[1] - cy) * scale];
        var fp1 = [cx + (p1[0] - cx) * scale, cy + (p1[1] - cy) * scale];
        var fp2 = [cx + (p2[0] - cx) * scale, cy + (p2[1] - cy) * scale];

        var whiteArrow = doc.pathItems.add();
        whiteArrow.setEntirePath([fp0, fp1, fp2]);
        whiteArrow.closed = true;
        whiteArrow.filled = true;
        whiteArrow.fillColor = whiteColor;
        whiteArrow.stroked = false;
        return whiteArrow;
    }

    function assembleLeaderParts(container, whitePath, newPath, whiteCircle, circle, hasWhiteEdge) {
        var edgeGrp = null;
        var mainGrp = container;

        if (hasWhiteEdge) {
            edgeGrp = container.groupItems.add();
            mainGrp = container.groupItems.add();

            if (newPath) newPath.move(mainGrp, ElementPlacement.PLACEATEND);
            if (circle) circle.move(mainGrp, ElementPlacement.PLACEATEND);

            if (whitePath) whitePath.move(edgeGrp, ElementPlacement.PLACEATEND);
            if (whiteCircle) whiteCircle.move(edgeGrp, ElementPlacement.PLACEATEND);
        } else {
            if (newPath) newPath.move(container, ElementPlacement.PLACEATEND);
            if (circle) circle.move(container, ElementPlacement.PLACEATEND);
        }

        return {
            edgeGroup: edgeGrp,
            mainGroup: mainGrp
        };
    }

    function tagLeaderParts(parts, whitePath, newPath, whiteCircle, circle, hasWhiteEdge) {
        if (hasWhiteEdge) {
            safeSetNote(parts.mainGroup, "leader_line_main");
            safeSetNote(parts.edgeGroup, "leader_line_edge");
        }

        safeSetNote(newPath, "leader_line_main_path");
        safeSetNote(circle, "leader_line_main_cap");
        safeSetNote(whitePath, "leader_line_edge_path");
        safeSetNote(whiteCircle, "leader_line_edge_cap");
    }

    function setLeaderLineTag(item) {
        safeSetNote(item, "leader_line");
    }

    function setTagValue(item, name, value) {
        if (!item) return;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) {
                    item.tags[i].value = String(value);
                    return;
                }
            }
            var t = item.tags.add();
            t.name = name;
            t.value = String(value);
        } catch (e) { }
    }

    function getTagValue(item, name) {
        if (!item) return null;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) return item.tags[i].value;
            }
        } catch (e) { }
        return null;
    }

    function setLeaderLineBoundsTags(item, b) {
        if (!item || !b) return;
        setTagValue(item, "leader_line_x_left", formatNumber(b.x_left, 4));
        setTagValue(item, "leader_line_x_right", formatNumber(b.x_right, 4));
        setTagValue(item, "leader_line_y_top", formatNumber(b.y_top, 4));
        setTagValue(item, "leader_line_y_bottom", formatNumber(b.y_bottom, 4));
    }

    function setLeaderLineDirTags(item, hDir, vDir) {
        if (!item) return;
        setTagValue(item, "leader_line_hDir", hDir);
        setTagValue(item, "leader_line_vDir", vDir);
    }

    function getLeaderLineStoredDir(item) {
        var hDir = getTagValue(item, "leader_line_hDir");
        var vDir = getTagValue(item, "leader_line_vDir");
        if (!hDir || !vDir) return null;
        return { hDir: hDir, vDir: vDir };
    }

    function getLeaderLineStoredBounds(item) {
        var xLeft = parseFloat(getTagValue(item, "leader_line_x_left"));
        var xRight = parseFloat(getTagValue(item, "leader_line_x_right"));
        var yTop = parseFloat(getTagValue(item, "leader_line_y_top"));
        var yBottom = parseFloat(getTagValue(item, "leader_line_y_bottom"));
        if (isNaN(xLeft) || isNaN(xRight) || isNaN(yTop) || isNaN(yBottom)) return null;
        return {
            x_left: xLeft,
            x_right: xRight,
            y_top: yTop,
            y_bottom: yBottom
        };
    }

    function isLeaderLineGroup(item) {
        return item && item.typename === "GroupItem" && item.note === "leader_line";
    }

    // (getOrCreateLeaderLineBackupLayer, stashOriginalTarget, getTargetCreationLayer removed)

    function isWhiteCMYKColor(color) {
        if (!color) return false;
        if (color.typename !== "CMYKColor") return false;
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }

    function getLeaderLineBasePath(groupItem) {
        var i, pi;

        // 再適用時の基準線は常に A線のみとし、B/C/D は参照しない
        // 新構造なら note 付きの A線を最優先
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.note === "leader_line_main_path") return pi;
        }

        // 旧構造用フォールバック：
        // 白でない stroked の open path を優先
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.stroked && !pi.closed && !isWhiteCMYKColor(pi.strokeColor)) return pi;
        }

        // 最後のフォールバック：stroked の open path
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.stroked && !pi.closed) return pi;
        }

        return null;
    }

    // プレビューを生成する関数
    function createPreview(angleDeg, ui) {
        var angleRad = angleDeg * Math.PI / 180;
        var boundsData = collectBoundsData(targets);
        var diagDirPreview = getDiagDirValues(ui);
        var hDir = diagDirPreview.hDir;
        var vDir = diagDirPreview.vDir;
        var addCircle = ui.capCircle.value;
        var addArrow = ui.capArrow.value;
        var addCap = addCircle || addArrow;
        var diameter = unitValueToPt(ui.capSizeInput.text, "strokeUnits");
        if (isNaN(diameter) || diameter <= 0) diameter = 3;

        for (var i = 0; i < boundsData.length; i++) {
            var b = boundsData[i];

            // 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用
            var objHDir = hDir;
            var objVDir = vDir;
            if (ui.scopeExceptDir.value) {
                var storedDir = getLeaderLineStoredDir(targets[i]);
                if (storedDir) {
                    objHDir = storedDir.hDir;
                    objVDir = storedDir.vDir;
                }
            }

            var sw = getLineWidth(b);
            var points = calcLeaderPoints(b, angleRad, objHDir, objVDir);

            // 先端座標は短縮前に取得
            var origTipPt = addCap ? getTipPoint(points, objHDir).slice(0) : null;
            var bendPt = addArrow ? points[1].slice(0) : null;

            if (addCircle && !ui.capFill.value) {
                // 円の半径 + 円の線幅の半分で短縮（線が円の縁に接する）
                var shortenR = diameter / 2 + sw / 2;
                shortenTip(points, objHDir, shortenR);
            }
            if (addArrow) {
                // 矢印の長さ分だけ短縮
                shortenTip(points, objHDir, diameter);
            }

            var whitePath = null;
            var newPath = null;
            var whiteCircle = null;
            var circle = null;

            // フチ（最背面）
            if (ui.whiteEdgeCheck.value) {
                whitePath = doc.pathItems.add();
                whitePath.setEntirePath(points);
                whitePath.stroked = true;
                whitePath.strokeWidth = sw * 3;
                whitePath.strokeColor = getEdgeColor();
                whitePath.filled = false;
                whitePath.strokeCap = ui.capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
            }

            newPath = doc.pathItems.add();
            newPath.setEntirePath(points);
            newPath.stroked = true;
            newPath.strokeWidth = sw;
            newPath.strokeColor = getLineColor(b);
            newPath.filled = false;
            newPath.strokeCap = ui.capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            if (addCircle) {
                if (ui.whiteEdgeCheck.value) {
                    whiteCircle = addTipCircleWhiteEdge(origTipPt, b, diameter, ui.capFill.value);
                }
                circle = addTipCircle(origTipPt, b, diameter, ui.capFill.value);
            }
            if (addArrow) {
                if (ui.whiteEdgeCheck.value) {
                    whiteCircle = addTipArrowWhiteEdge(origTipPt, bendPt, b, diameter, ui.capFill.value);
                }
                circle = addTipArrow(origTipPt, bendPt, b, diameter, ui.capFill.value);
            }

            var previewGrp = doc.groupItems.add();
            assembleLeaderParts(previewGrp, whitePath, newPath, whiteCircle, circle, ui.whiteEdgeCheck.value); previewItems.push(previewGrp);
        }
    }

    // プレビューを削除する関数
    function removePreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            safeRemove(previewItems[i]);
        }
        previewItems = [];
    }

    function setTargetsVisible(visible) {
        for (var i = 0; i < targets.length; i++) {
            targets[i].hidden = !visible;
        }
    }

    // ↑↓キーで値を増減する関数
    function changeValueByArrowKey(editText, options) {
        options = options || {};
        var smallStep = (typeof options.smallStep === "number") ? options.smallStep : 1;
        var largeStep = (typeof options.largeStep === "number") ? options.largeStep : 10;
        var fineStep = (typeof options.fineStep === "number") ? options.fineStep : 0.1;
        var minValue = (typeof options.minValue === "number") ? options.minValue : 0;
        var digits = (typeof options.digits === "number") ? options.digits : 0;
        var onAfterChange = (typeof options.onAfterChange === "function") ? options.onAfterChange : null;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = parseFloat(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var step = smallStep;
            if (keyboard.shiftKey) {
                step = largeStep;
            } else if (keyboard.altKey) {
                step = fineStep;
            }

            if (keyboard.shiftKey) {
                // Shift: snap to next/prev multiple of largeStep
                if (event.keyName === "Up") {
                    var ceil = Math.ceil(value / step * (1 + 1e-9)) * step;
                    value = (ceil <= value) ? value + step : ceil;
                } else {
                    var floor = Math.floor(value / step * (1 - 1e-9)) * step;
                    value = (floor >= value) ? value - step : floor;
                    if (value < minValue) value = minValue;
                }
            } else if (event.keyName === "Up") {
                value += step;
            } else {
                value -= step;
                if (value < minValue) value = minValue;
            }

            editText.text = formatNumber(value, digits);
            event.preventDefault();
            if (onAfterChange) onAfterChange();
        });
    }

    // 前回の設定を取得 / Load previous session settings
    var s = $.global._leaderLineSettings;
    var strokeUnitLabel = getCurrentUnitLabel("strokeUnits");
    var __zoomState = __TMKZoom_captureViewState(doc);

    function loadDialogSettingsToUI(ui) {
        ui.angleInput.text = s.angle;
        ui.radio30.value = (s.radioAngle === 30);
        ui.radio45.value = (s.radioAngle !== 30 && s.radioAngle !== 60);
        ui.radio60.value = (s.radioAngle === 60);

        if (!s.hasUserSetApplyScope && targets.length > 1) {
            ui.scopeExceptDir.value = true;
            ui.scopeAll.value = false;
        } else if (s.applyScope === "exceptDirection") {
            ui.scopeExceptDir.value = true;
            ui.scopeAll.value = false;
        } else {
            ui.scopeAll.value = true;
            ui.scopeExceptDir.value = false;
        }

        if (s.diagDir === "lowerLeft") {
            ui.dirLowerLeft.value = true;
            ui.dirUpperLeft.value = false;
            ui.dirUpperRight.value = false;
            ui.dirLowerRight.value = false;
        } else if (s.diagDir === "upperRight") {
            ui.dirUpperRight.value = true;
            ui.dirUpperLeft.value = false;
            ui.dirLowerLeft.value = false;
            ui.dirLowerRight.value = false;
        } else if (s.diagDir === "lowerRight") {
            ui.dirLowerRight.value = true;
            ui.dirUpperLeft.value = false;
            ui.dirLowerLeft.value = false;
            ui.dirUpperRight.value = false;
        } else {
            ui.dirUpperLeft.value = true;
            ui.dirLowerLeft.value = false;
            ui.dirUpperRight.value = false;
            ui.dirLowerRight.value = false;
        }

        if (s.lineColor === "white") {
            ui.colorWhite.value = true;
            ui.colorBlack.value = false;
            ui.colorOther.value = false;
        } else if (s.lineColor === "other") {
            ui.colorOther.value = true;
            ui.colorBlack.value = false;
            ui.colorWhite.value = false;
        } else {
            ui.colorBlack.value = true;
            ui.colorWhite.value = false;
            ui.colorOther.value = false;
        }
        ui.hexInput.text = s.lineColorHex;
        updateSwatch();

        lineWidthDisplay = parseFloat(formatUnitValue(s.lineWidth, "strokeUnits"));
        if (isNaN(lineWidthDisplay) || lineWidthDisplay <= 0) lineWidthDisplay = 1;
        ui.lineWidthInput.text = formatNumber(lineWidthDisplay, 2);

        if (s.strokeCapType === "round") {
            ui.capRound.value = true;
            ui.strokeCapNone.value = false;
        } else {
            ui.strokeCapNone.value = true;
            ui.capRound.value = false;
        }

        if (s.capType === "circle") {
            ui.capCircle.value = true;
            ui.capNone.value = false;
            ui.capArrow.value = false;
        } else if (s.capType === "arrow") {
            ui.capArrow.value = true;
            ui.capNone.value = false;
            ui.capCircle.value = false;
        } else {
            ui.capNone.value = true;
            ui.capCircle.value = false;
            ui.capArrow.value = false;
        }

        if (s.capStyle === "stroke") {
            ui.capStroke.value = true;
            ui.capFill.value = false;
        } else {
            ui.capFill.value = true;
            ui.capStroke.value = false;
        }

        capSizeDisplay = parseFloat(formatUnitValue(s.capSize, "strokeUnits"));
        if (isNaN(capSizeDisplay) || capSizeDisplay <= 0) capSizeDisplay = 3;
        ui.capSizeInput.text = formatNumber(capSizeDisplay, 2);

        ui.groupCheck.value = !!s.groupEnabled;
        ui.whiteEdgeCheck.value = !!s.whiteEdge;

        if (s.edgeColor === "other") {
            ui.edgeColorOther.value = true;
            ui.edgeColorWhite.value = false;
        } else {
            ui.edgeColorWhite.value = true;
            ui.edgeColorOther.value = false;
        }
        ui.edgeHexInput.text = s.edgeColorHex;
        updateEdgeSwatch();

        updateDirPanelEnabled(ui);
        updateCapEnabled(ui);
        updateEdgeColorEnabled(ui);
    }


    function buildDialogUI() {
        var ui = {};

        ui.dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        ui.dlg.orientation = "column";
        ui.dlg.alignChildren = ["fill", "top"];
        // ダイアログ表示位置は復元する
        if (s.dialogBounds && s.dialogBounds.length === 2) {
            ui.dlg.location = s.dialogBounds;
        }

        // 適用範囲ラジオボタン
        ui.scopePanel = ui.dlg.add("panel", undefined, L("panelApplyScope"));
        ui.scopePanel.margins = [15, 20, 15, 10];
        ui.scopePanel.orientation = "row";
        ui.scopePanel.alignChildren = ["center", "center"];
        ui.scopeAll = ui.scopePanel.add("radiobutton", undefined, L("scopeAll"));
        ui.scopeExceptDir = ui.scopePanel.add("radiobutton", undefined, L("scopeExceptDirection"));

        ui.topRow = ui.dlg.add("group");
        ui.topRow.orientation = "row";
        ui.topRow.alignChildren = ["fill", "top"];

        ui.leftCol = ui.topRow.add("group");
        ui.leftCol.orientation = "column";
        ui.leftCol.alignChildren = ["fill", "top"];

        ui.anglePanel = ui.leftCol.add("panel", undefined, L("panelAngle"));
        ui.anglePanel.margins = [15, 20, 15, 10];
        ui.anglePanel.orientation = "column";
        ui.anglePanel.alignChildren = ["fill", "top"];

        ui.angleInputGroup = ui.anglePanel.add("group");
        ui.angleInputGroup.alignment = ["center", "top"];
        ui.angleInput = ui.angleInputGroup.add("edittext", undefined, s.angle);
        ui.angleInput.characters = 4;
        ui.angleInputGroup.add("statictext", undefined, "\u00B0");
        ui.angleInput.active = true;

        ui.radioGroup = ui.anglePanel.add("group");
        ui.radioGroup.orientation = "row";
        ui.radio30 = ui.radioGroup.add("radiobutton", undefined, "30");
        ui.radio45 = ui.radioGroup.add("radiobutton", undefined, "45");
        ui.radio60 = ui.radioGroup.add("radiobutton", undefined, "60");

        ui.dirPanel = ui.leftCol.add("panel", undefined, L("panelDirection"));
        ui.dirPanel.margins = [15, 20, 15, 10];
        ui.dirPanel.orientation = "column";
        ui.dirPanel.alignChildren = ["fill", "top"];

        ui.dirGroup = ui.dirPanel.add("group");
        ui.dirGroup.orientation = "row";
        ui.dirGroup.alignChildren = ["fill", "top"];

        ui.dirLeftCol = ui.dirGroup.add("group");
        ui.dirLeftCol.orientation = "column";
        ui.dirLeftCol.alignChildren = ["fill", "top"];
        ui.dirUpperLeft = ui.dirLeftCol.add("radiobutton", undefined, L("dirUpperLeft"));
        ui.dirLeftCol.add("statictext", undefined, " ");
        ui.dirLowerLeft = ui.dirLeftCol.add("radiobutton", undefined, L("dirLowerLeft"));

        ui.dirCenterCol = ui.dirGroup.add("group");
        ui.dirCenterCol.orientation = "column";
        ui.dirCenterCol.alignChildren = ["center", "center"];
        ui.dirCenterCol.add("statictext", undefined, " ");
        ui.dirCenterCol.add("statictext", undefined, "対象");

        ui.dirRightCol = ui.dirGroup.add("group");
        ui.dirRightCol.orientation = "column";
        ui.dirRightCol.alignChildren = ["fill", "top"];
        ui.dirUpperRight = ui.dirRightCol.add("radiobutton", undefined, L("dirUpperRight"));
        ui.dirRightCol.add("statictext", undefined, " ");
        ui.dirLowerRight = ui.dirRightCol.add("radiobutton", undefined, L("dirLowerRight"));

        ui.rightCol = ui.topRow.add("group");
        ui.rightCol.orientation = "column";
        ui.rightCol.alignChildren = ["fill", "top"];

        ui.colorPanel = ui.rightCol.add("panel", undefined, L("panelLineStyle"));
        ui.colorPanel.margins = [15, 20, 15, 10];
        ui.colorPanel.orientation = "column";
        ui.colorPanel.alignChildren = ["fill", "top"];

        ui.colorGroup = ui.colorPanel.add("group");
        ui.colorGroup.alignChildren = ["left", "center"];
        ui.colorBlack = ui.colorGroup.add("radiobutton", undefined, L("colorBlack"));
        ui.colorWhite = ui.colorGroup.add("radiobutton", undefined, L("colorWhite"));
        ui.colorOther = ui.colorGroup.add("radiobutton", undefined, "");
        ui.hexInput = { text: s.lineColorHex };
        ui.colorSwatch = ui.colorGroup.add("panel", undefined, "");
        ui.colorSwatch.preferredSize = [20, 20];

        ui.lineWidthGroup = ui.colorPanel.add("group");
        ui.lineWidthGroup.add("statictext", undefined, L("lineWidth"));
        lineWidthDisplay = parseFloat(formatUnitValue(s.lineWidth, "strokeUnits"));
        if (isNaN(lineWidthDisplay) || lineWidthDisplay <= 0) lineWidthDisplay = 1;
        ui.lineWidthInput = ui.lineWidthGroup.add("edittext", undefined, formatNumber(lineWidthDisplay, 2));
        ui.lineWidthInput.characters = 3;
        ui.lineWidthGroup.add("statictext", undefined, strokeUnitLabel);

        ui.strokeCapGroup = ui.colorPanel.add("group");
        ui.strokeCapNone = ui.strokeCapGroup.add("radiobutton", undefined, L("capStrokeCapNone"));
        ui.capRound = ui.strokeCapGroup.add("radiobutton", undefined, L("capRound"));

        ui.capPanel = ui.rightCol.add("panel", undefined, L("panelLineEnd"));
        ui.capPanel.margins = [15, 20, 15, 10];
        ui.capPanel.orientation = "column";
        ui.capPanel.alignChildren = ["fill", "top"];

        ui.capGroup = ui.capPanel.add("group");
        ui.capNone = ui.capGroup.add("radiobutton", undefined, L("capNone"));
        ui.capCircle = ui.capGroup.add("radiobutton", undefined, L("capCircle"));
        ui.capArrow = ui.capGroup.add("radiobutton", undefined, L("capArrow"));

        ui.capStyleGroup = ui.capPanel.add("group");
        ui.capFill = ui.capStyleGroup.add("radiobutton", undefined, L("capFill"));
        ui.capStroke = ui.capStyleGroup.add("radiobutton", undefined, L("capStroke"));

        ui.capSizeGroup = ui.capPanel.add("group");
        ui.capSizeGroup.add("statictext", undefined, L("capSize"));
        capSizeDisplay = parseFloat(formatUnitValue(s.capSize, "strokeUnits"));
        if (isNaN(capSizeDisplay) || capSizeDisplay <= 0) capSizeDisplay = 3;
        ui.capSizeInput = ui.capSizeGroup.add("edittext", undefined, formatNumber(capSizeDisplay, 2));
        ui.capSizeInput.characters = 3;
        ui.capSizeGroup.add("statictext", undefined, strokeUnitLabel);

        ui.groupCheck = ui.capPanel.add("checkbox", undefined, L("groupEnabled"));

        ui.optPanel = ui.leftCol.add("panel", undefined, L("panelOptions"));
        ui.optPanel.margins = [15, 20, 15, 10];
        ui.optPanel.orientation = "column";
        ui.optPanel.alignChildren = ["fill", "top"];

        ui.whiteEdgeCheck = ui.optPanel.add("checkbox", undefined, L("whiteEdge"));
        ui.edgeColorGroup = ui.optPanel.add("group");
        ui.edgeColorGroup.orientation = "row";
        ui.edgeColorGroup.alignChildren = ["left", "center"];
        ui.edgeColorWhite = ui.edgeColorGroup.add("radiobutton", undefined, L("colorWhite"));
        ui.edgeColorOther = ui.edgeColorGroup.add("radiobutton", undefined, "");
        ui.edgeHexInput = { text: s.edgeColorHex };
        ui.edgeSwatch = ui.edgeColorGroup.add("panel", undefined, "");
        ui.edgeSwatch.preferredSize = [20, 20];

        return ui;
    }

    function saveDialogSettingsFromUI(ui) {
        s.angle = ui.angleInput.text;
        s.radioAngle = ui.radio30.value ? 30 : (ui.radio60.value ? 60 : 45);
        s.applyScope = ui.scopeExceptDir.value ? "exceptDirection" : "all";

        var diagDirVals = getDiagDirValues(ui);
        s.diagDir = ui.dirUpperLeft.value ? "upperLeft" : (ui.dirLowerLeft.value ? "lowerLeft" : (ui.dirUpperRight.value ? "upperRight" : "lowerRight"));
        s.hDir = diagDirVals.hDir;
        s.vDir = diagDirVals.vDir;

        s.capType = ui.capCircle.value ? "circle" : (ui.capArrow.value ? "arrow" : "none");
        s.capStyle = ui.capStroke.value ? "stroke" : "fill";
        var savedCapSizePt = unitValueToPt(ui.capSizeInput.text, "strokeUnits");
        s.capSize = (!isNaN(savedCapSizePt) && savedCapSizePt > 0) ? formatNumber(savedCapSizePt, 4) : "3";
        s.strokeCapType = ui.capRound.value ? "round" : "none";

        s.groupEnabled = ui.groupCheck.value;
        s.whiteEdge = ui.whiteEdgeCheck.value;
        s.edgeColor = ui.edgeColorWhite.value ? "white" : "other";
        s.edgeColorHex = ui.edgeHexInput.text;

        s.lineColor = ui.colorBlack.value ? "black" : (ui.colorWhite.value ? "white" : "other");
        s.lineColorHex = ui.hexInput.text;
        var savedLineWidthPt = unitValueToPt(ui.lineWidthInput.text, "strokeUnits");
        s.lineWidth = (!isNaN(savedLineWidthPt) && savedLineWidthPt > 0) ? formatNumber(savedLineWidthPt, 4) : "1";

        rememberDialogLocation();
    }

    var ui = buildDialogUI();
    var dlg = ui.dlg;

    var scopeAll = ui.scopeAll;
    var scopeExceptDir = ui.scopeExceptDir;

    var dirPanel = ui.dirPanel;

    var colorBlack = ui.colorBlack;
    var colorWhite = ui.colorWhite;
    var colorOther = ui.colorOther;
    var hexInput = ui.hexInput;
    var colorSwatch = ui.colorSwatch;

    var lineWidthInput = ui.lineWidthInput;

    var strokeCapNone = ui.strokeCapNone;
    var capRound = ui.capRound;

    var capNone = ui.capNone;
    var capCircle = ui.capCircle;
    var capArrow = ui.capArrow;

    var capFill = ui.capFill;
    var capStroke = ui.capStroke;
    var capSizeInput = ui.capSizeInput;

    var groupCheck = ui.groupCheck;

    var whiteEdgeCheck = ui.whiteEdgeCheck;
    var edgeColorGroup = ui.edgeColorGroup;
    var edgeColorWhite = ui.edgeColorWhite;
    var edgeColorOther = ui.edgeColorOther;
    var edgeHexInput = ui.edgeHexInput;
    var edgeSwatch = ui.edgeSwatch;

    // ズーム状態は保存しないが、ダイアログ位置は保存する
    function rememberDialogLocation() {
        if (dlg.location) {
            s.dialogBounds = [dlg.location[0], dlg.location[1]];
        }
    }

    function getDiagDirValues(ui) {
        if (ui.dirUpperLeft.value) return { hDir: "right", vDir: "down" };
        if (ui.dirLowerLeft.value) return { hDir: "right", vDir: "up" };
        if (ui.dirUpperRight.value) return { hDir: "left", vDir: "down" };
        if (ui.dirLowerRight.value) return { hDir: "left", vDir: "up" };
        return { hDir: "right", vDir: "up" };
    }

    function updateDirPanelEnabled(ui) {
        ui.dirPanel.enabled = ui.scopeAll.value;
    }

    function updateCapEnabled(ui) {
        var on = ui.capCircle.value || ui.capArrow.value;
        if (ui.capArrow.value) {
            if (ui.capStroke.value) {
                ui.capStroke.value = false;
                ui.capFill.value = true;
            }
            ui.capFill.enabled = false;
            ui.capStroke.enabled = false;
        } else {
            ui.capFill.enabled = on;
            ui.capStroke.enabled = on;
        }
        ui.capSizeInput.enabled = on;
        ui.groupCheck.enabled = on;
    }

    function updateEdgeColorEnabled(ui) {
        var on = ui.whiteEdgeCheck.value;
        ui.edgeColorGroup.enabled = on;
        ui.edgeSwatch.enabled = on;
    }

    function bindDialogEvents(ui) {
        changeValueByArrowKey(ui.angleInput, {
            smallStep: 1,
            largeStep: 10,
            fineStep: 0.1,
            minValue: 0,
            digits: 0,
            onAfterChange: function () { updatePreview(ui); }
        });
        changeValueByArrowKey(ui.lineWidthInput, {
            smallStep: 0.1,
            largeStep: 1,
            fineStep: 0.01,
            minValue: 0,
            digits: 2,
            onAfterChange: function () { updatePreview(ui); }
        });
        changeValueByArrowKey(ui.capSizeInput, {
            smallStep: 0.1,
            largeStep: 1,
            fineStep: 0.01,
            minValue: 0,
            digits: 2,
            onAfterChange: function () { updatePreview(ui); }
        });

        ui.radio30.onClick = function () { ui.angleInput.text = "30"; updatePreview(ui); };
        ui.radio45.onClick = function () { ui.angleInput.text = "45"; updatePreview(ui); };
        ui.radio60.onClick = function () { ui.angleInput.text = "60"; updatePreview(ui); };

        // 適用範囲はユーザーが明示的にクリックしたときだけ「手動設定済み」とみなす
        ui.scopeAll.onClick = function () {
            s.hasUserSetApplyScope = true;
            updateDirPanelEnabled(ui);
            updatePreview(ui);
        };
        ui.scopeExceptDir.onClick = function () {
            s.hasUserSetApplyScope = true;
            updateDirPanelEnabled(ui);
            updatePreview(ui);
        };

        var allDirRadios = [ui.dirUpperLeft, ui.dirLowerLeft, ui.dirUpperRight, ui.dirLowerRight];
        function uncheckOtherDirs(selected) {
            for (var d = 0; d < allDirRadios.length; d++) {
                if (allDirRadios[d] !== selected) allDirRadios[d].value = false;
            }
        }
        ui.dirUpperLeft.onClick = function () { uncheckOtherDirs(ui.dirUpperLeft); updatePreview(ui); };
        ui.dirLowerLeft.onClick = function () { uncheckOtherDirs(ui.dirLowerLeft); updatePreview(ui); };
        ui.dirUpperRight.onClick = function () { uncheckOtherDirs(ui.dirUpperRight); updatePreview(ui); };
        ui.dirLowerRight.onClick = function () { uncheckOtherDirs(ui.dirLowerRight); updatePreview(ui); };

        ui.colorSwatch.addEventListener("click", function () {
            var initial = ui.hexInput.text.replace(/^#/, "");
            var c = ColorPicker.show(initial);
            if (c) {
                ui.hexInput.text = "#" + c;
                ui.colorBlack.value = false;
                ui.colorWhite.value = false;
                ui.colorOther.value = true;
                updateSwatch();
                updatePreview(ui);
            }
        });
        ui.colorBlack.onClick = function () { updatePreview(ui); };
        ui.colorWhite.onClick = function () { updatePreview(ui); };
        ui.colorOther.onClick = function () {
            var initial = ui.hexInput.text.replace(/^#/, "");
            var c = ColorPicker.show(initial);
            if (c) {
                ui.hexInput.text = "#" + c;
                updateSwatch();
            }
            updatePreview(ui);
        };
        ui.lineWidthInput.onChanging = function () { updatePreview(ui); };

        ui.capRound.onClick = function () { updatePreview(ui); };
        ui.strokeCapNone.onClick = function () { updatePreview(ui); };
        ui.capNone.onClick = function () { updateCapEnabled(ui); updatePreview(ui); };
        ui.capCircle.onClick = function () { updateCapEnabled(ui); updatePreview(ui); };
        ui.capArrow.onClick = function () { updateCapEnabled(ui); updatePreview(ui); };
        ui.capFill.onClick = function () { updatePreview(ui); };
        ui.capStroke.onClick = function () { updatePreview(ui); };
        ui.capSizeInput.onChanging = function () { updatePreview(ui); };
        ui.groupCheck.onClick = function () { saveDialogSettingsFromUI(ui); };

        ui.whiteEdgeCheck.onClick = function () {
            updateEdgeColorEnabled(ui);
            updatePreview(ui);
        };
        ui.edgeSwatch.addEventListener("click", function () {
            var initial = ui.edgeHexInput.text.replace(/^#/, "");
            var c = ColorPicker.show(initial);
            if (c) {
                ui.edgeHexInput.text = "#" + c;
                ui.edgeColorWhite.value = false;
                ui.edgeColorOther.value = true;
                updateEdgeSwatch();
                updatePreview(ui);
            }
        });
        ui.edgeColorWhite.onClick = function () { updatePreview(ui); };
        ui.edgeColorOther.onClick = function () {
            var initial = ui.edgeHexInput.text.replace(/^#/, "");
            var c = ColorPicker.show(initial);
            if (c) {
                ui.edgeHexInput.text = "#" + c;
                updateEdgeSwatch();
            }
            updatePreview(ui);
        };

        ui.angleInput.onChanging = function () { updatePreview(ui); };

        ui.dlg.onMove = function () {
            rememberDialogLocation();
        };
        ui.dlg.onClose = function () {
            saveDialogSettingsFromUI(ui);
        };
    }

    bindDialogEvents(ui);
    loadDialogSettingsToUI(ui);

    var zoomCtrl = __TMKZoom_addControls(dlg, doc, L("zoom"), __zoomState, {
        min: 0.1,
        max: 8,
        sliderWidth: 240,
        margins: [0, 0, 0, 10],
        redraw: true,
        lightMode: true,
        lightModeLabel: L("lightMode"),
        lightModeDefault: false
    });
    ui.zoomCtrl = zoomCtrl;

    // ズーム状態はセッション復元しない。ここでは操作時の挙動だけ定義する
    if (zoomCtrl.lightModeCheckbox) {
        zoomCtrl.lightModeCheckbox.onClick = function () {
            zoomCtrl.applyZoom(Number(zoomCtrl.slider.value));
        };
    }

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["center", "top"];
    btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    function updatePreview(ui) {
        removePreview();
        setTargetsVisible(true);

        var val = parseFloat(ui.angleInput.text);
        if (isNaN(val) || val <= 0 || val >= 90) {
            saveDialogSettingsFromUI(ui);
            app.redraw();
            return;
        }

        createPreview(val, ui);
        setTargetsVisible(false);
        saveDialogSettingsFromUI(ui);
        app.redraw();
    }

    updatePreview(ui);

    var result = dlg.show();
    saveDialogSettingsFromUI(ui);

    // プレビューを削除して元のパスを復元
    removePreview();
    setTargetsVisible(true);
    app.redraw();

    if (result !== 1) {
        zoomCtrl.restoreInitial();
        return;
    }

    var angleDeg = parseFloat(ui.angleInput.text);
    if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
        alert(L("alertInvalidAngle"));
        app.redraw();
        return;
    }
    var angleRad = angleDeg * Math.PI / 180;

    // 確定：引き出し線を生成

    var diagDirResult = getDiagDirValues(ui);
    var hDir = diagDirResult.hDir;
    var vDir = diagDirResult.vDir;
    var addCircle = ui.capCircle.value;
    var addArrow = ui.capArrow.value;
    var addCap = addCircle || addArrow;
    var diameter = unitValueToPt(ui.capSizeInput.text, "strokeUnits");
    if (isNaN(diameter) || diameter <= 0) diameter = 3;
    var boundsData = collectBoundsData(targets);
    var finalSelectionItems = [];
    doc.selection = null;

    for (var i = 0; i < targets.length; i++) {
        var sourceTarget = targets[i];
        var b = boundsData[i];
        var createdItems = [];
        var createdSelectionItems = [];

        try {
            // 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用
            var objHDir = hDir;
            var objVDir = vDir;
            if (ui.scopeExceptDir.value) {
                var storedDir = getLeaderLineStoredDir(sourceTarget);
                if (storedDir) {
                    objHDir = storedDir.hDir;
                    objVDir = storedDir.vDir;
                }
            }

            var sw = getLineWidth(b);
            var points = calcLeaderPoints(b, angleRad, objHDir, objVDir);

            // 先端座標は短縮前に取得
            var origTipPt = addCap ? getTipPoint(points, objHDir).slice(0) : null;
            var bendPt = addArrow ? points[1].slice(0) : null;

            if (addCircle && !ui.capFill.value) {
                var shortenR = diameter / 2 + sw / 2;
                shortenTip(points, objHDir, shortenR);
            }
            if (addArrow) {
                shortenTip(points, objHDir, diameter);
            }

            var whitePath = null;
            var newPath = null;
            var whiteCircle = null;
            var circle = null;

            // フチ（最背面）
            if (ui.whiteEdgeCheck.value) {
                whitePath = doc.pathItems.add();
                createdItems.push(whitePath);
                whitePath.setEntirePath(points);
                whitePath.stroked = true;
                whitePath.strokeWidth = sw * 3;
                whitePath.strokeColor = getEdgeColor();
                whitePath.filled = false;
                whitePath.strokeCap = ui.capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            }

            newPath = doc.pathItems.add();
            createdItems.push(newPath);
            newPath.setEntirePath(points);
            newPath.stroked = true;
            newPath.strokeWidth = sw;
            newPath.strokeColor = getLineColor(b);
            newPath.filled = false;
            newPath.strokeCap = ui.capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            if (addCircle) {
                if (ui.whiteEdgeCheck.value) {
                    whiteCircle = addTipCircleWhiteEdge(origTipPt, b, diameter, ui.capFill.value);
                }
                circle = addTipCircle(origTipPt, b, diameter, ui.capFill.value);
                createdItems.push(circle);
            }
            if (addArrow) {
                if (ui.whiteEdgeCheck.value) {
                    whiteCircle = addTipArrowWhiteEdge(origTipPt, bendPt, b, diameter, ui.capFill.value);
                }
                circle = addTipArrow(origTipPt, bendPt, b, diameter, ui.capFill.value);
                createdItems.push(circle);
            }

            if (ui.groupCheck.value) {
                var grp = doc.groupItems.add();
                createdItems.push(grp);
                var assembled = assembleLeaderParts(grp, whitePath, newPath, whiteCircle, circle, ui.whiteEdgeCheck.value);
                tagLeaderParts(assembled, whitePath, newPath, whiteCircle, circle, ui.whiteEdgeCheck.value);
                setLeaderLineTag(grp);
                setLeaderLineBoundsTags(grp, b);
                setLeaderLineDirTags(grp, objHDir, objVDir);
                createdSelectionItems.push(grp);
            } else {
                if (whitePath) createdSelectionItems.push(whitePath);
                if (whiteCircle) createdSelectionItems.push(whiteCircle);
                if (newPath) createdSelectionItems.push(newPath);
                if (circle) createdSelectionItems.push(circle);
            }

            for (var cs = 0; cs < createdSelectionItems.length; cs++) {
                finalSelectionItems.push(createdSelectionItems[cs]);
            }

            // 置き換え成功時のみ元オブジェクトを削除
            sourceTarget.remove();
        } catch (eTarget) {
            for (var cr = createdItems.length - 1; cr >= 0; cr--) {
                safeRemove(createdItems[cr]);
            }
            throw eTarget;
        }
    }

    if (finalSelectionItems.length) {
        doc.selection = null;
        for (var si = 0; si < finalSelectionItems.length; si++) {
            safeSelect(finalSelectionItems[si]);
        }
    }
}

main();