#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
【概要 / Summary】
複数アートボードのドキュメントで、選択中のオブジェクトを整列（移動）します。

■ 対象
- アートボード：選択オブジェクトを「中心点が属するアートボード」ごとに振り分け、各アートボード内の指定位置へ整列します。
- 選択オブジェクト：アクティブなアートボード上の選択を“テンプレ”として位置関係（オフセット）を取得し、他のアートボード上で選択しているオブジェクトを同じ位置関係になるように整列します（アクティブ側は動かしません）。

■ 整列先
- 角：左上／右上／左下／右下
- 中央：アートボード中央
※「中央」選択時はマージンは無効（0扱い）です。
※「選択オブジェクト」対象時は「中央」は無効（ディム表示）です。
※「選択オブジェクト」対象時も、この「整列先」を基準にテンプレの相対位置を複製します。

■ マージン（左右／上下）
- 角に整列する際、角から内側へオフセットできます（左右・上下を別指定）
- 「連動」ON のときは、左右の入力が操作元になり、上下はディム表示で同値に追従します（デフォルト：ON）
- マージンの単位は Illustrator のルーラー単位（rulerType）に連動し、入力値はその単位として解釈して内部では pt に変換します。
- 「対象：選択オブジェクト」選択時はマージンは無視（0扱い）になります。

■ プレビューと操作
- ダイアログ表示中はプレビューで動作を確認でき、［OK］で確定します（キャンセル／×では元に戻ります）
- 整列先はキーボードでも選択できます：w=左上 / e=右上 / s=左下 / d=右下 / c=中央
- 対象はキーボードでも選択できます：a=アートボード / q=選択オブジェクト
- 「対象：選択オブジェクト」選択時、アクティブなアートボード上のテンプレ選択が上半分なら「左上」、下半分なら「右下」を自動選択します（初回のみ。以後は手動で変更できます）。
- マージン欄は ↑↓ で増減（Shift=±10、Option=±0.1）し、入力中もプレビューが更新されます
- Enter/Return キーで［OK］を実行できます

更新日 / Updated: 2025-12-17
*/

var SCRIPT_NAME = "Artboard Corner Align";
var SCRIPT_VERSION = "v1.1";
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
    target: { ja: "対象", en: "Target" },
    targetArtboard: { ja: "アートボード", en: "Artboard" },
    targetSelection: { ja: "選択オブジェクト", en: "Selection" },
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
    invalidMargin: { ja: "マージンは数値で指定してください。", en: "Margin must be a number." }
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
    var targetMode = input.targetMode || "artboard"; // "artboard" | "selection"
    var margin = input.marginPt; // pt: number or {x,y}

    if (targetMode === "selection") {
        // Template: use selection position in ACTIVE artboard, then apply same relative offset to other artboards
        var abIndexActive = -1;
        try { abIndexActive = doc.artboards.getActiveArtboardIndex(); } catch (eAb) { abIndexActive = -1; }
        if (abIndexActive < 0) return;

        var mapAll = groupSelectedItemsByArtboard(doc, doc.selection);
        if (!mapAll) return;

        var refItems = mapAll[abIndexActive];
        if (!refItems || refItems.length === 0) return; // must have template on active artboard

        var refRect = doc.artboards[abIndexActive].artboardRect;
        var offset = computeGroupOffsetFromArtboard(refItems, refRect, corner);
        if (!offset) return;

        for (var abIndex in mapAll) {
            if (!mapAll.hasOwnProperty(abIndex)) continue;
            var abi = parseInt(abIndex, 10);
            if (abi === abIndexActive) continue; // active artboard is the template; do not move

            var items = mapAll[abIndex];
            if (!items || items.length === 0) continue;

            var rect = doc.artboards[abi].artboardRect;
            alignGroupToArtboardWithOffset(items, rect, corner, offset);
        }

        try { app.redraw(); } catch (eRedrawMain1) {}

    } else {
        // Default: distribute per artboard based on each item's center
        var map = groupSelectedItemsByArtboard(doc, doc.selection);
        if (!map) return;

        for (var abIndex in map) {
            if (!map.hasOwnProperty(abIndex)) continue;

            var items = map[abIndex];
            if (!items || items.length === 0) continue;

            var rect = doc.artboards[parseInt(abIndex, 10)].artboardRect; // [L, T, R, B]
            alignItemsToCornerWithMargin(items, rect, corner, margin);
        }
        try { app.redraw(); } catch (eRedrawMain2) {}
    }
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

    function applyPreview(corner, marginX, marginY, targetMode) {
        // Always revert first to avoid cumulative moves
        revertPreview(false);

        if (!doc || !previewState.items || previewState.items.length === 0) return;

        if (targetMode === "selection") {
            var abIndexActive = -1;
            try { abIndexActive = doc.artboards.getActiveArtboardIndex(); } catch (eAb2) { abIndexActive = -1; }
            if (abIndexActive < 0) return;

            var mapAll = groupSelectedItemsByArtboard(doc, previewState.items);
            if (!mapAll) return;

            var refItems = mapAll[abIndexActive];
            if (!refItems || refItems.length === 0) return;

            var refRect = doc.artboards[abIndexActive].artboardRect;
            var offset = computeGroupOffsetFromArtboard(refItems, refRect, corner);
            if (!offset) return;

            for (var abIndex in mapAll) {
                if (!mapAll.hasOwnProperty(abIndex)) continue;
                var abi = parseInt(abIndex, 10);
                if (abi === abIndexActive) continue;

                var items = mapAll[abIndex];
                if (!items || items.length === 0) continue;

                var rect = doc.artboards[abi].artboardRect;
                alignGroupToArtboardWithOffsetPreview(items, rect, corner, offset, previewState);
            }

        } else {
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
    // Layout
    // ----------------------------
    // 上：2カラム（左＝整列先 / 右＝マージン）
    // 下：OK/キャンセル（左右に並べる）
    w.orientation = "column";
    w.alignChildren = ["fill", "top"];

    var gTop = w.add("group");
    gTop.orientation = "row";
    gTop.alignChildren = ["fill", "top"];

    // 2カラム構成：左＝整列先、右＝マージン
    var gLeft = gTop.add("group");
    gLeft.orientation = "column";
    gLeft.alignChildren = ["fill", "top"];

    var gRight = gTop.add("group");
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
    // Target shortcut keys: A = Artboard, Q = Selection
    w.addEventListener("keydown", function (event) {
        var k = event.keyName;
        if (!k) return;
        k = ("" + k).toLowerCase();

        if (k === "a") {
            rbTargetArtboard.value = true;
            event.preventDefault();
            updatePreviewFromUI();
        } else if (k === "q") {
            rbTargetSelection.value = true;
            event.preventDefault();
            updatePreviewFromUI();
        }
    });

    // --- Target Panel (logic TBD)
    var pTarget = gRight.add("panel", undefined, _(LABELS.target.ja, LABELS.target.en));
    pTarget.orientation = "column";
    pTarget.alignChildren = ["left", "top"];
    pTarget.margins = [15, 20, 15, 10];

    var rbTargetArtboard = pTarget.add("radiobutton", undefined, _(LABELS.targetArtboard.ja, LABELS.targetArtboard.en));
    var rbTargetSelection = pTarget.add("radiobutton", undefined, _(LABELS.targetSelection.ja, LABELS.targetSelection.en));
    rbTargetSelection.value = true;

    // Helper: check if target is "Selection"
    function isTargetSelection() {
        try { return rbTargetSelection.value === true; } catch (e) { return false; }
    }

    // --- Auto-corner selection flag and helper ---
    var _autoCornerUpdating = false;
    var _autoCornerAppliedOnce = false;
    var _lastTargetIsSelection = null;

    function getAutoCornerFromActiveTemplate() {
        // Returns "LT" for upper half, "RB" for lower half, or null if cannot determine
        try {
            if (!doc) return null;
            var abIndexActive = -1;
            try { abIndexActive = doc.artboards.getActiveArtboardIndex(); } catch (eAb) { abIndexActive = -1; }
            if (abIndexActive < 0) return null;

            // Use current (reverted) previewState items to match preview behavior
            var mapAll = groupSelectedItemsByArtboard(doc, previewState.items);
            if (!mapAll) return null;

            var refItems = mapAll[abIndexActive];
            if (!refItems || refItems.length === 0) return null;

            var refRect = doc.artboards[abIndexActive].artboardRect; // [L,T,R,B]
            var gb = getUnionVisibleBounds(refItems); // [L,T,R,B]
            if (!gb) return null;

            var abCy = (refRect[1] + refRect[3]) / 2;
            var gCy = (gb[1] + gb[3]) / 2;

            // Illustrator coordinate: larger Y is higher. So gCy > abCy means upper half.
            return (gCy >= abCy) ? "LT" : "RB";
        } catch (eAuto) {
            return null;
        }
    }

    // ラジオボタン変更でプレビュー更新 / Update preview on radio change
    rbLT.onClick = updatePreviewFromUI;
    rbRT.onClick = updatePreviewFromUI;
    rbLB.onClick = updatePreviewFromUI;
    rbRB.onClick = updatePreviewFromUI;
    rbC.onClick = updatePreviewFromUI;
    // Target radio buttons also trigger preview/UI update
    rbTargetArtboard.onClick = updatePreviewFromUI;
    rbTargetSelection.onClick = updatePreviewFromUI;

    // --- Margin Panel: 2-column layout inside the panel
    var pMargin = gRight.add("panel", undefined, _(LABELS.marginValue.ja, LABELS.marginValue.en) + " (" + unitLabel + ")");
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
    changeValueByArrowKey(etMarginX, true, function () {
        if (cbLink.value) {
            etMarginY.text = etMarginX.text;
        }
        updatePreviewFromUI();
    });
    changeValueByArrowKey(etMarginY, true, function () {
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
    // --- Buttons (bottom, centered, side-by-side)
    var gButtons = w.add("group");
    gButtons.orientation = "row";
    gButtons.alignChildren = ["center", "center"];
    gButtons.alignment = ["center", "bottom"];

    var btnCancel = gButtons.add("button", undefined, _(LABELS.cancel.ja, LABELS.cancel.en), { name: "cancel" });
    var btnOK = gButtons.add("button", undefined, _(LABELS.ok.ja, LABELS.ok.en), { name: "ok" });

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

    // Margin panel enable/disable based on "中央" and Target radio
    function updatePreviewFromUI() {
        var corner = getCornerFromUI();
        var isCenter = (corner === "C");
        var targetIsSelection = isTargetSelection();

        // Auto-select corner based on template position when Target=Selection (only once on entering selection mode)
        if (_lastTargetIsSelection === null) _lastTargetIsSelection = targetIsSelection;

        if (targetIsSelection && !_lastTargetIsSelection) {
            // Entering selection mode: reset and apply auto-corner once
            _autoCornerAppliedOnce = false;
        }
        if (!targetIsSelection && _lastTargetIsSelection) {
            // Leaving selection mode: reset for next time
            _autoCornerAppliedOnce = false;
        }

        if (targetIsSelection && !_autoCornerAppliedOnce && !_autoCornerUpdating) {
            var autoCorner = getAutoCornerFromActiveTemplate();
            if (autoCorner === "LT" || autoCorner === "RB") {
                _autoCornerUpdating = true;
                try {
                    rbLT.value = (autoCorner === "LT");
                    rbRT.value = false;
                    rbLB.value = false;
                    rbRB.value = (autoCorner === "RB");
                    rbC.value = false;
                } catch (eSetAuto) {}
                _autoCornerUpdating = false;

                corner = autoCorner;
                isCenter = false;
                _autoCornerAppliedOnce = true;
            }
        }

        _lastTargetIsSelection = targetIsSelection;

        // Disable Center option when Target=Selection
        try {
            rbC.enabled = !targetIsSelection;
            if (targetIsSelection && rbC.value) {
                rbLT.value = true; // fallback to LT
                corner = "LT";
                isCenter = false;
            }
        } catch (eCenterDim) {}

        // Disable margin when Center is selected OR when Target = Selection
        var enableMargin = (!isCenter && !targetIsSelection);
        try { pMargin.enabled = enableMargin; } catch (eDim) {}
        // Also dim link checkbox
        try { cbLink.enabled = enableMargin; } catch (eDim2) {}

        var mxInput = safeToNumber(etMarginX.text);
        var myInput = safeToNumber(etMarginY.text);
        if (mxInput === null) mxInput = 0;
        if (myInput === null) myInput = 0;

        // In selection mode, margin is ignored
        if (targetIsSelection) {
            mxInput = 0;
            myInput = 0;
        }

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
            if (mxPt === null) mxPt = 0;
            if (myPt === null) myPt = 0;
        }
        applyPreview(corner, mxPt, myPt, targetIsSelection ? "selection" : "artboard");
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
        var targetMode = isTargetSelection() ? "selection" : "artboard";
        var mxInput = safeToNumber(etMarginX.text);
        var myInput = safeToNumber(etMarginY.text);
        if (mxInput === null) mxInput = 0;
        if (myInput === null) myInput = 0;
        if (!isCenter && targetMode !== "selection") {
            if (cbLink.value) {
                if (mxInput === null) {
                    alert(_(LABELS.invalidMargin.ja, LABELS.invalidMargin.en));
                    return;
                }
            } else {
                if (mxInput === null || myInput === null) {
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
    var mxPt = 0, myPt = 0;
    if (!isCenter) {
        if (cbLink.value) {
            mxPt = toPt(mxInput, unitLabel);
            myPt = toPt(mxInput, unitLabel);
            if (mxPt === null) return null;
            if (myPt === null) return null;
        } else {
            mxPt = toPt(mxInput, unitLabel);
            myPt = toPt(myInput, unitLabel);
            if (mxPt === null) return null;
            if (myPt === null) return null;
        }
    }
    var targetMode = isTargetSelection() ? "selection" : "artboard";
    return {
        corner: corner,
        targetMode: targetMode,
        // If Target=Selection, margin is ignored (treated as 0)
        marginPt: (isCenter || targetMode === "selection") ? 0 : { x: mxPt, y: myPt },
        marginXPt: (targetMode === "selection") ? 0 : mxPt,
        marginYPt: (targetMode === "selection") ? 0 : myPt,
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

function filterItemsInArtboard(doc, items, artboardIndex) {
    var out = [];
    if (!items || items.length === 0) return out;

    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        if (isLockedOrHidden(it)) continue;

        var vb;
        try { vb = it.visibleBounds; } catch (e) { continue; }

        var cx = (vb[0] + vb[2]) / 2;
        var cy = (vb[1] + vb[3]) / 2;

        var ab = findArtboardIndexByPoint(doc, cx, cy);
        if (ab === artboardIndex) out.push(it);
    }
    return out;
}

function getUnionVisibleBounds(items) {
    if (!items || items.length === 0) return null;

    var l = null, t = null, r = null, b = null;
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;

        var vb;
        try { vb = it.visibleBounds; } catch (e) { continue; }

        if (l === null) {
            l = vb[0]; t = vb[1]; r = vb[2]; b = vb[3];
        } else {
            if (vb[0] < l) l = vb[0];
            if (vb[1] > t) t = vb[1];
            if (vb[2] > r) r = vb[2];
            if (vb[3] < b) b = vb[3];
        }
    }

    if (l === null) return null;
    return [l, t, r, b];
}

function alignSelectionGroupToCorner(items, artboardRect, corner) {
    var gb = getUnionVisibleBounds(items);
    if (!gb) return;

    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    var gL = gb[0], gT = gb[1], gR = gb[2], gB = gb[3];
    var gCx = (gL + gR) / 2;
    var gCy = (gT + gB) / 2;

    var dx = 0, dy = 0;
    if (corner === "C") {
        dx = abCx - gCx;
        dy = abCy - gCy;
    } else if (corner === "LT") {
        dx = abL - gL;
        dy = abT - gT;
    } else if (corner === "RT") {
        dx = abR - gR;
        dy = abT - gT;
    } else if (corner === "LB") {
        dx = abL - gL;
        dy = abB - gB;
    } else { // RB
        dx = abR - gR;
        dy = abB - gB;
    }

    for (var i = 0; i < items.length; i++) {
        try { items[i].translate(dx, dy); } catch (e) {}
    }
}

function alignSelectionGroupToCornerPreview(items, artboardRect, corner, previewState) {
    var gb = getUnionVisibleBounds(items);
    if (!gb) return;

    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    var gL = gb[0], gT = gb[1], gR = gb[2], gB = gb[3];
    var gCx = (gL + gR) / 2;
    var gCy = (gT + gB) / 2;

    var dx = 0, dy = 0;
    if (corner === "C") {
        dx = abCx - gCx;
        dy = abCy - gCy;
    } else if (corner === "LT") {
        dx = abL - gL;
        dy = abT - gT;
    } else if (corner === "RT") {
        dx = abR - gR;
        dy = abT - gT;
    } else if (corner === "LB") {
        dx = abL - gL;
        dy = abB - gB;
    } else { // RB
        dx = abR - gR;
        dy = abB - gB;
    }

    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;

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

function getArtboardAnchorPoint(artboardRect, corner) {
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    if (corner === "C") return [abCx, abCy];
    if (corner === "LT") return [abL, abT];
    if (corner === "RT") return [abR, abT];
    if (corner === "LB") return [abL, abB];
    return [abR, abB]; // RB
}

function getBoundsAnchorPoint(bounds, corner) {
    // bounds: [L,T,R,B]
    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3];
    var cx = (l + r) / 2;
    var cy = (t + b) / 2;

    if (corner === "C") return [cx, cy];
    if (corner === "LT") return [l, t];
    if (corner === "RT") return [r, t];
    if (corner === "LB") return [l, b];
    return [r, b]; // RB
}

function computeGroupOffsetFromArtboard(items, artboardRect, corner) {
    // Returns {x,y} offset = (groupAnchor - artboardAnchor)
    var gb = getUnionVisibleBounds(items);
    if (!gb) return null;

    var abA = getArtboardAnchorPoint(artboardRect, corner);
    var gA = getBoundsAnchorPoint(gb, corner);

    return { x: (gA[0] - abA[0]), y: (gA[1] - abA[1]) };
}

function alignGroupToArtboardWithOffset(items, artboardRect, corner, offset) {
    // Moves all items together so that their groupAnchor equals (artboardAnchor + offset)
    var gb = getUnionVisibleBounds(items);
    if (!gb) return;

    var abA = getArtboardAnchorPoint(artboardRect, corner);
    var targetX = abA[0] + (offset.x || 0);
    var targetY = abA[1] + (offset.y || 0);

    var gA = getBoundsAnchorPoint(gb, corner);
    var dx = targetX - gA[0];
    var dy = targetY - gA[1];

    for (var i = 0; i < items.length; i++) {
        try { items[i].translate(dx, dy); } catch (e) {}
    }
}

function alignGroupToArtboardWithOffsetPreview(items, artboardRect, corner, offset, previewState) {
    var gb = getUnionVisibleBounds(items);
    if (!gb) return;

    var abA = getArtboardAnchorPoint(artboardRect, corner);
    var targetX = abA[0] + (offset.x || 0);
    var targetY = abA[1] + (offset.y || 0);

    var gA = getBoundsAnchorPoint(gb, corner);
    var dx = targetX - gA[0];
    var dy = targetY - gA[1];

    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;

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