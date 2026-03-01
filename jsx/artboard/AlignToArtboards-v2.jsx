#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
【概要 / Summary】
複数アートボードのドキュメントで、選択中のオブジェクトを「中心点が属するアートボード」ごとに振り分け、
各アートボード内の指定位置へ一括で整列（移動）します。

■ 整列先
- 角：左上／右上／左下／右下
- 中央：アートボード中央
※「中央」選択時はマージンは無効（0扱い）です。

■ マージン（左右／上下）
- 角に整列する際、角から内側へオフセットできます（左右・上下を別指定）
- 「連動」ON のときは、左右の入力が操作元になり、上下はディム表示で同値に追従します（デフォルト：ON）
- マージンの単位は Illustrator のルーラー単位（rulerType）に連動し、入力値はその単位として解釈して内部では pt に変換します。

■ プレビューと操作
- ダイアログ表示中はプレビューで動作を確認でき、［OK］で確定します（キャンセル／×では元に戻ります）
- 角はキーボードでも選択できます：w=左上 / e=右上 / s=左下 / d=右下 / c=中央
- マージン欄は ↑↓ で増減（Shift=±10、Option=±0.1）し、入力中もプレビューが更新されます
- Enter/Return キーで［OK］を実行できます

更新日 / Updated: 2025-12-17
*/

var SCRIPT_NAME = "Artboard Corner Align";
var SCRIPT_VERSION = "v1.0";
var UPDATED_DATE = "2025-12-17";

// --------------------------------------
// ローカライズ / Localization
// --------------------------------------
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// Simple label switcher
function _(ja, en) {
    return (lang === "ja") ? ja : en;
}

var LABELS = {
    dialogTitle: { ja: "各アートボードに整列 " + SCRIPT_VERSION, en: "Align to Artboards" + SCRIPT_VERSION },
    targetCorner: { ja: "整列先", en: "Target" },
    leftTop: { ja: "左上", en: "Top-Left" },
    rightTop: { ja: "右上", en: "Top-Right" },
    leftBottom: { ja: "左下", en: "Bottom-Left" },
    rightBottom: { ja: "右下", en: "Bottom-Right" },
    center: { ja: "中央", en: "Center" },
    marginValue: { ja: "マージン", en: "Margin" },
    marginX: { ja: "左右", en: "Horizontal" },
    marginY: { ja: "上下", en: "Vertical" },
    link: { ja: "連動", en: "Linked" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    invalidMargin: { ja: "マージンは 0 以上の数値で指定してください。", en: "Margin must be a number >= 0." }
};

// --------------------------------------
// 単位（rulerType） / Units (rulerType)
// --------------------------------------
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

function getCurrentUnitLabel() {
    var unitCode = 2; // pt
    try { unitCode = app.preferences.getIntegerPreference("rulerType"); } catch (e) {}
    return unitLabelMap[unitCode] || "pt";
}

// Convert number (typed in current unit) to points.
function toPt(valueNumber, unitLabel) {
    if (valueNumber === null || valueNumber === undefined) return null;
    if (isNaN(valueNumber)) return null;

    // Normalize a few labels to UnitValue-compatible units
    var u = unitLabel;
    if (u === "pica") u = "pc";
    else if (u !== "in" && u !== "mm" && u !== "cm" && u !== "pt" && u !== "px") {
        // Unknown/rare unit: treat as pt (safe fallback)
        u = "pt";
    }

    try {
        return new UnitValue(valueNumber, u).as("pt");
    } catch (eUV) {
        return valueNumber; // fallback: assume pt
    }
}

// --------------------------------------
// メイン / Main
// --------------------------------------
function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) return;

    var input = showDialog();
    if (input === null) return;

    var corner = input.corner;   // "LT" "RT" "LB" "RB" "C"
    var margin = input.marginPt; // number (pt)

    var map = groupSelectedItemsByArtboard(doc, doc.selection);
    if (!map) return;

    for (var abIndex in map) {
        if (!map.hasOwnProperty(abIndex)) continue;

        var items = map[abIndex];
        if (!items || items.length === 0) continue;

        var rect = doc.artboards[parseInt(abIndex, 10)].artboardRect; // [L, T, R, B]
        alignItemsToCornerWithMargin(items, rect, corner, margin);
    }
    try { app.redraw(); } catch (eRedrawMain) {}
}

main();

// --------------------------------------
// UI / Dialog
// --------------------------------------
function showDialog() {
    var w = new Window("dialog", (lang === "ja") ? LABELS.dialogTitle.ja : LABELS.dialogTitle.en);
    var unitLabel = getCurrentUnitLabel();

    // --------------------------------------
    // プレビュー状態 / Preview state (moves are reverted when dialog closes)
    // --------------------------------------
    var doc = app.activeDocument;
    var previewState = {
        items: [],
        dx: [],
        dy: [],
        applied: false
    };
    var closingByOK = false; // OK 経由で閉じるときの redraw 抑止

    try {
        // Keep references to current selection for preview
        if (doc && doc.selection && doc.selection.length > 0) {
            for (var pi = 0; pi < doc.selection.length; pi++) {
                previewState.items.push(doc.selection[pi]);
                previewState.dx.push(0);
                previewState.dy.push(0);
            }
        }
    } catch (ePrevInit) {}

    function revertPreview(doRedraw) {
        if (doRedraw === undefined) doRedraw = true;
        if (!previewState.applied) return;
        for (var i = 0; i < previewState.items.length; i++) {
            var it = previewState.items[i];
            var dx = previewState.dx[i];
            var dy = previewState.dy[i];
            if (!it) continue;
            try {
                if (dx !== 0 || dy !== 0) it.translate(-dx, -dy);
            } catch (eRev) {}
            previewState.dx[i] = 0;
            previewState.dy[i] = 0;
        }
        previewState.applied = false;
        if (doRedraw) {
            try { app.redraw(); } catch (eRedraw1) {}
        }
    }

    function applyPreview(corner, marginX, marginY) {
        // Always revert first to avoid cumulative moves
        revertPreview(false);

        if (!doc || !previewState.items || previewState.items.length === 0) return;

        // Group by artboard based on the (reverted) current positions
        var map = groupSelectedItemsByArtboard(doc, previewState.items);
        if (!map) return;

        // Note: marginX and marginY are in pt
        for (var abIndex in map) {
            if (!map.hasOwnProperty(abIndex)) continue;
            var items = map[abIndex];
            if (!items || items.length === 0) continue;

            var rect = doc.artboards[parseInt(abIndex, 10)].artboardRect; // [L, T, R, B]
            alignItemsToCornerWithMarginPreviewXY(items, rect, corner, marginX, marginY, previewState);
        }

        previewState.applied = true;
        try { app.redraw(); } catch (eRedraw2) {}
    }

    function getCornerFromUI() {
        if (rbRT.value) return "RT";
        if (rbLB.value) return "LB";
        if (rbRB.value) return "RB";
        if (rbC.value) return "C";
        return "LT";
    }

    function addCornerKeyHandler(dialog, rbLT, rbRT, rbLB, rbRB, rbC) {
        dialog.addEventListener("keydown", function (event) {
            // Key names are case-insensitive in practice, normalize.
            var k = event.keyName;
            if (!k) return;
            k = ("" + k).toLowerCase();

            if (k === "w") {
                rbLT.value = true;
                event.preventDefault();
                updatePreviewFromUI();
            } else if (k === "e") {
                rbRT.value = true;
                event.preventDefault();
                updatePreviewFromUI();
            } else if (k === "s") {
                rbLB.value = true;
                event.preventDefault();
                updatePreviewFromUI();
            } else if (k === "d") {
                rbRB.value = true;
                event.preventDefault();
                updatePreviewFromUI();
            } else if (k === "c") {
                rbC.value = true;
                event.preventDefault();
                updatePreviewFromUI();
            }
        });
    }

    // ----------------------------
    // Margin Panel with Link
    // ----------------------------
    w.orientation = "row";
    w.alignChildren = ["fill", "top"];

    // 2カラム構成：左＝設定、右＝ボタン
    var gLeft = w.add("group");
    gLeft.orientation = "column";
    gLeft.alignChildren = ["fill", "top"];

    var gRight = w.add("group");
    gRight.orientation = "column";
    gRight.alignChildren = ["fill", "top"];
    gRight.alignment = ["right", "top"];

    var pCorner = gLeft.add("panel", undefined, _(LABELS.targetCorner.ja, LABELS.targetCorner.en));
    pCorner.orientation = "column";
    pCorner.alignChildren = ["left", "top"];
    pCorner.margins = [15, 20, 15, 10];

    var rbLT = pCorner.add("radiobutton", undefined, _(LABELS.leftTop.ja, LABELS.leftTop.en));
    var rbRT = pCorner.add("radiobutton", undefined, _(LABELS.rightTop.ja, LABELS.rightTop.en));
    var rbLB = pCorner.add("radiobutton", undefined, _(LABELS.leftBottom.ja, LABELS.leftBottom.en));
    var rbRB = pCorner.add("radiobutton", undefined, _(LABELS.rightBottom.ja, LABELS.rightBottom.en));
    var rbC = pCorner.add("radiobutton", undefined, _(LABELS.center.ja, LABELS.center.en));
    rbLT.value = true;
    addCornerKeyHandler(w, rbLT, rbRT, rbLB, rbRB, rbC);

    // ラジオボタン変更でプレビュー更新 / Update preview on radio change
    rbLT.onClick = updatePreviewFromUI;
    rbRT.onClick = updatePreviewFromUI;
    rbLB.onClick = updatePreviewFromUI;
    rbRB.onClick = updatePreviewFromUI;
    rbC.onClick = updatePreviewFromUI;

    // --- Margin Panel: 2-column layout inside the panel
    var pMargin = gLeft.add("panel", undefined, _(LABELS.marginValue.ja, LABELS.marginValue.en) + " (" + unitLabel + ")");
    pMargin.orientation = "row";
    pMargin.alignChildren = ["fill", "center"];
    pMargin.margins = [15, 20, 15, 10];

    var leftColumn = pMargin.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["left", "center"];
    var rightColumn = pMargin.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["left", "center"];

    // Horizontal (左右) group
    var gMX = leftColumn.add("group");
    gMX.orientation = "row";
    gMX.alignChildren = ["left", "center"];
    var stMarginX = gMX.add("statictext", undefined, _(LABELS.marginX.ja, LABELS.marginX.en));
    var etMarginX = gMX.add("edittext", undefined, "0");
    etMarginX.characters = 4;

    // Vertical (上下) group
    var gMY = leftColumn.add("group");
    gMY.orientation = "row";
    gMY.alignChildren = ["left", "center"];
    var stMarginY = gMY.add("statictext", undefined, _(LABELS.marginY.ja, LABELS.marginY.en));
    var etMarginY = gMY.add("edittext", undefined, "0");
    etMarginY.characters = 4;

    // Link checkbox (right column)
    var cbLink = rightColumn.add("checkbox", undefined, _(LABELS.link.ja, LABELS.link.en));
    cbLink.value = true;

    // Arrow key helpers and change events
    changeValueByArrowKey(etMarginX, false, function () {
        if (cbLink.value) {
            etMarginY.text = etMarginX.text;
        }
        updatePreviewFromUI();
    });
    changeValueByArrowKey(etMarginY, false, function () {
        if (!cbLink.value) updatePreviewFromUI();
    });

    etMarginX.onChange = function () {
        if (cbLink.value) {
            etMarginY.text = etMarginX.text;
        }
        updatePreviewFromUI();
    };
    etMarginX.onChanging = function () {
        if (cbLink.value) {
            etMarginY.text = etMarginX.text;
        }
        updatePreviewFromUI();
    };

    etMarginY.onChange = function () {
        if (!cbLink.value) updatePreviewFromUI();
    };
    etMarginY.onChanging = function () {
        if (!cbLink.value) updatePreviewFromUI();
    };

    // Enter-to-OK for both fields
    function bindEnterToOK(editText) {
        editText.addEventListener("keydown", function (ev) {
            if (ev.keyName === "Enter" || ev.keyName === "Return") {
                try { btnOK.notify(); } catch (eNotify) {}
                ev.preventDefault();
            }
        });
    }
    var btnOK = gRight.add("button", undefined, _(LABELS.ok.ja, LABELS.ok.en), { name: "ok" });
    var btnCancel = gRight.add("button", undefined, _(LABELS.cancel.ja, LABELS.cancel.en), { name: "cancel" });
    try { w.defaultElement = btnOK; } catch (eDef) {}
    bindEnterToOK(etMarginX);
    bindEnterToOK(etMarginY);

    // Link checkbox behavior
    function updateLinkStateUI() {
        if (cbLink.value) {
            // Linked ON: 左右を操作元に、上下をディム
            etMarginX.enabled = true;
            etMarginY.enabled = false;
            etMarginY.text = etMarginX.text;
        } else {
            // Linked OFF: 両方操作可
            etMarginX.enabled = true;
            etMarginY.enabled = true;
        }
        updatePreviewFromUI();
    }
    cbLink.onClick = updateLinkStateUI;

    // Margin panel enable/disable based on "中央"
    function updatePreviewFromUI() {
        var corner = getCornerFromUI();
        var isCenter = (corner === "C");
        try { pMargin.enabled = !isCenter; } catch (eDim) {}
        // Also dim link checkbox
        try { cbLink.enabled = !isCenter; } catch (eDim2) {}

        var mxInput = safeToNumber(etMarginX.text);
        var myInput = safeToNumber(etMarginY.text);
        if (mxInput === null) mxInput = 0;
        if (myInput === null) myInput = 0;
        if (mxInput < 0) mxInput = 0;
        if (myInput < 0) myInput = 0;
        var mxPt = 0, myPt = 0;
        if (!isCenter) {
            if (cbLink.value) {
                // Linked: use X for both
                mxPt = toPt(mxInput, unitLabel);
                myPt = toPt(mxInput, unitLabel);
            } else {
                mxPt = toPt(mxInput, unitLabel);
                myPt = toPt(myInput, unitLabel);
            }
            if (mxPt === null || mxPt < 0) mxPt = 0;
            if (myPt === null || myPt < 0) myPt = 0;
        }
        applyPreview(corner, mxPt, myPt);
    }

    // Dialog close: always revert preview (OK後は main() で確定実行)
    w.onClose = function () {
        revertPreview(closingByOK ? false : true);
        return true;
    };

    // ダイアログ表示時：マージン（左右）をフォーカス＆ハイライト
    w.onShow = function () {
        try {
            etMarginX.active = true;
            etMarginX.selection = [0, (etMarginX.text || "").length];
        } catch (eFocus) {}
    };

    // 初期表示でプレビュー適用（選択がある場合）
    updateLinkStateUI();

    btnOK.onClick = function () {
        var corner = getCornerFromUI();
        var isCenter = (corner === "C");
        var mxInput = safeToNumber(etMarginX.text);
        var myInput = safeToNumber(etMarginY.text);
        if (mxInput === null) mxInput = 0;
        if (myInput === null) myInput = 0;
        if (mxInput < 0) mxInput = 0;
        if (myInput < 0) myInput = 0;
        if (!isCenter) {
            if (cbLink.value) {
                if (mxInput === null || mxInput < 0) {
                    alert(_(LABELS.invalidMargin.ja, LABELS.invalidMargin.en));
                    return;
                }
            } else {
                if (mxInput === null || mxInput < 0 || myInput === null || myInput < 0) {
                    alert(_(LABELS.invalidMargin.ja, LABELS.invalidMargin.en));
                    return;
                }
            }
        }
        closingByOK = true;
        revertPreview(false);
        w.close(1);
    };

    if (w.show() !== 1) return null;

    var corner = getCornerFromUI();
    var isCenter = (corner === "C");
    var mxInput = safeToNumber(etMarginX.text);
    var myInput = safeToNumber(etMarginY.text);
    if (mxInput === null) mxInput = 0;
    if (myInput === null) myInput = 0;
    if (mxInput < 0) mxInput = 0;
    if (myInput < 0) myInput = 0;
    var mxPt = 0, myPt = 0;
    if (!isCenter) {
        if (cbLink.value) {
            mxPt = toPt(mxInput, unitLabel);
            myPt = toPt(mxInput, unitLabel);
            if (mxPt === null || mxPt < 0) return null;
            if (myPt === null || myPt < 0) return null;
        } else {
            mxPt = toPt(mxInput, unitLabel);
            myPt = toPt(myInput, unitLabel);
            if (mxPt === null || mxPt < 0) return null;
            if (myPt === null || myPt < 0) return null;
        }
    }
    // Return: if linked and not center, marginX = marginY
    return {
        corner: corner,
        marginPt: isCenter ? 0 : { x: mxPt, y: myPt },
        marginXPt: mxPt,
        marginYPt: myPt,
        linked: cbLink.value
    };
}

function safeToNumber(s) {
    if (s === null || s === undefined) return null;
    s = ("" + s).replace(/^\s+|\s+$/g, "");
    if (s === "") return null;
    var n = Number(s);
    if (isNaN(n)) return null;
    return n;
}

function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
    if (allowNegative === undefined) allowNegative = false;

    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
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

        // 丸め
        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; // 小数第1位まで
        } else {
            value = Math.round(value); // 整数
        }

        if (!allowNegative && value < 0) value = 0;

        editText.text = value;

        // プレビュー更新など（任意）
        if (typeof onValueChanged === "function") {
            try { onValueChanged(); } catch (eCb) {}
        }

        event.preventDefault();
    });
}

// --------------------------------------
// 選択をアートボードへ振り分け / Group selection by artboard
// --------------------------------------
function groupSelectedItemsByArtboard(doc, selection) {
    var result = {};

    for (var i = 0; i < selection.length; i++) {
        var it = selection[i];
        if (isLockedOrHidden(it)) continue;

        var vb;
        try { vb = it.visibleBounds; } catch (e) { continue; } // [L,T,R,B]

        var cx = (vb[0] + vb[2]) / 2;
        var cy = (vb[1] + vb[3]) / 2;

        var abIndex = findArtboardIndexByPoint(doc, cx, cy);
        if (abIndex === -1) continue;

        if (!result[abIndex]) result[abIndex] = [];
        result[abIndex].push(it);
    }

    return result;
}

function findArtboardIndexByPoint(doc, x, y) {
    var n = doc.artboards.length;
    for (var i = 0; i < n; i++) {
        var r = doc.artboards[i].artboardRect; // [L, T, R, B]
        if (x >= r[0] && x <= r[2] && y <= r[1] && y >= r[3]) return i;
    }
    return -1;
}

function isLockedOrHidden(item) {
    try { if (item.locked) return true; } catch (e) {}
    try { if (item.hidden) return true; } catch (e2) {}

    var p = item.parent;
    while (p) {
        try { if (p.locked) return true; } catch (e3) {}
        try { if (p.hidden) return true; } catch (e4) {}
        if (p.typename === "Document") break;
        p = p.parent;
    }
    return false;
}

// --------------------------------------
// 整列（移動） / Align (move) with margin
// --------------------------------------
function alignItemsToCornerWithMargin(items, artboardRect, corner, margin) {
    // margin: pt (number or {x, y})
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;
    var mx = 0, my = 0;
    if (typeof margin === "object" && margin !== null) {
        mx = margin.x || 0;
        my = margin.y || 0;
    } else {
        mx = my = margin || 0;
    }
    var targetX, targetY;
    if (corner === "LT") {
        targetX = abL + mx;
        targetY = abT - my;
    } else if (corner === "RT") {
        targetX = abR - mx;
        targetY = abT - my;
    } else if (corner === "LB") {
        targetX = abL + mx;
        targetY = abB + my;
    } else if (corner === "RB") {
        targetX = abR - mx;
        targetY = abB + my;
    }
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var vb;
        try { vb = it.visibleBounds; } catch (e) { continue; } // [L,T,R,B]
        var l = vb[0], t = vb[1], r = vb[2], b = vb[3];
        var dx = 0, dy = 0;
        if (corner === "C") {
            var cx = (l + r) / 2;
            var cy = (t + b) / 2;
            dx = abCx - cx;
            dy = abCy - cy;
        } else if (corner === "LT") {
            dx = targetX - l;
            dy = targetY - t;
        } else if (corner === "RT") {
            dx = targetX - r;
            dy = targetY - t;
        } else if (corner === "LB") {
            dx = targetX - l;
            dy = targetY - b;
        } else { // "RB"
            dx = targetX - r;
            dy = targetY - b;
        }
        try { it.translate(dx, dy); } catch (e2) {}
    }
}

function alignItemsToCornerWithMarginPreviewXY(items, artboardRect, corner, mx, my, previewState) {
    // mx, my: pt
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;
    var targetX, targetY;
    if (corner === "LT") {
        targetX = abL + mx;
        targetY = abT - my;
    } else if (corner === "RT") {
        targetX = abR - mx;
        targetY = abT - my;
    } else if (corner === "LB") {
        targetX = abL + mx;
        targetY = abB + my;
    } else if (corner === "RB") {
        targetX = abR - mx;
        targetY = abB + my;
    }
    // Record per-item delta so we can revert on close
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var vb;
        try { vb = it.visibleBounds; } catch (e) { continue; } // [L,T,R,B]
        var l = vb[0], t = vb[1], r = vb[2], b = vb[3];
        var dx = 0, dy = 0;
        if (corner === "C") {
            var cx = (l + r) / 2;
            var cy = (t + b) / 2;
            dx = abCx - cx;
            dy = abCy - cy;
        } else if (corner === "LT") {
            dx = targetX - l;
            dy = targetY - t;
        } else if (corner === "RT") {
            dx = targetX - r;
            dy = targetY - t;
        } else if (corner === "LB") {
            dx = targetX - l;
            dy = targetY - b;
        } else { // "RB"
            dx = targetX - r;
            dy = targetY - b;
        }
        // Find index in previewState.items
        var idx = -1;
        for (var j = 0; j < previewState.items.length; j++) {
            if (previewState.items[j] === it) { idx = j; break; }
        }
        try { it.translate(dx, dy); } catch (e2) {}
        if (idx >= 0) {
            previewState.dx[idx] += dx;
            previewState.dy[idx] += dy;
        }
    }
}