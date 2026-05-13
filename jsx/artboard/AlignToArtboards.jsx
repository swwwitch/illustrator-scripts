#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
【概要 / Summary】
複数アートボードのドキュメントで、選択中のオブジェクトを整列（移動）します。

### 整列の基準
- すべてのアートボード：選択オブジェクトを「中心点が属するアートボード」ごとに振り分け、各アートボード内の指定位置へ整列します。
- アクティブを基準：アクティブなアートボード上の選択を“テンプレ”として位置関係（オフセット）を取得し、他のアートボード上で選択しているオブジェクトを同じ位置関係になるように整列します（アクティブ側は動かしません）。

### 整列先
- 3×3 の9点：左上／上中央／右上／左中央／中央／右中央／左下／下中央／右下
※「中央」選択時はマージンは無効（0扱い）です。
※上中央／下中央は左右マージン無効、左中央／右中央は上下マージン無効です。
※「アクティブを基準」対象時は整列先パネル全体がディム表示になり、その時点の選択が維持されます。
※「アクティブを基準」対象時は、この「整列先」を基準にテンプレの相対位置を複製します。

### マージン（左右／上下）
- 整列する際、対応する辺から内側へオフセットできます（左右・上下を別指定）
- 「連動」ON のときは、左右の入力が操作元になり、上下はディム表示で同値に追従します（デフォルト：ON）
- マージンの単位は Illustrator のルーラー単位（rulerType）に連動し、入力値はその単位として解釈して内部では pt に変換します。
- 「整列の基準：アクティブを基準」選択時はマージンは無視（0扱い）になります。

### プレビューと操作
- ダイアログ表示中はプレビューで動作を確認でき、［OK］で確定します（キャンセル／×では元に戻ります）
- 整列先はキーボードでも選択できます：q=左上 / w=上中央 / e=右上 / a=左中央 / s=中央 / d=右中央 / z=左下 / x=下中央 / c=右下
- 整列の基準はキーボードでも選択できます：1=すべてのアートボード / 2=アクティブを基準
- マージン欄は ↑↓ で増減（Shift=±10、Option=±0.1）し、入力中もプレビューが更新されます
- Enter/Return キーで［OK］を実行できます

更新日 / Updated: 2025-12-17
*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.1";

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在のロケールを判定 / Detect current locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ロケールに応じてラベルを切替 / Switch label by locale */
function _(ja, en) {
    return (lang === "ja") ? ja : en;
}

var LABELS = {
    dialogTitle: { ja: "各アートボードに整列 " + SCRIPT_VERSION, en: "Align to Artboards" + SCRIPT_VERSION },
    targetCorner: { ja: "整列先", en: "Target" },
    target: { ja: "整列の基準", en: "Align based on" },
    targetArtboard: { ja: "すべてのアートボード", en: "All Artboards" },
    targetSelection: { ja: "アクティブを基準", en: "Based on Active" },
    leftTop: { ja: "左上", en: "Top-Left" },
    topCenter: { ja: "上中央", en: "Top-Center" },
    rightTop: { ja: "右上", en: "Top-Right" },
    leftMiddle: { ja: "左中央", en: "Middle-Left" },
    center: { ja: "中央", en: "Center" },
    rightMiddle: { ja: "右中央", en: "Middle-Right" },
    leftBottom: { ja: "左下", en: "Bottom-Left" },
    bottomCenter: { ja: "下中央", en: "Bottom-Center" },
    rightBottom: { ja: "右下", en: "Bottom-Right" },
    marginValue: { ja: "マージン", en: "Margin" },
    marginX: { ja: "左右", en: "Horizontal" },
    marginY: { ja: "上下", en: "Vertical" },
    link: { ja: "連動", en: "Linked" },
    usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};

// =========================================
// 境界取得（プレビュー境界 or 幾何境界） / Bounds resolver
// プレビュー境界＝ストローク・効果を含む、幾何境界＝図形のみ
// =========================================
var usePreviewBounds = true;

/* アイテムの整列用バウンディングボックスを返す / Return item bounds for alignment */
function getItemBounds(item) {
    try {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    } catch (e) {
        return null;
    }
}

// =========================================
// 単位（rulerType） / Units (rulerType)
// =========================================
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

/* 現在のルーラー単位ラベルを取得 / Get current ruler unit label */
function getCurrentUnitLabel() {
    var unitCode = 2; // pt
    try { unitCode = app.preferences.getIntegerPreference("rulerType"); } catch (e) { }
    return unitLabelMap[unitCode] || "pt";
}

/* 入力値（現在単位）を pt に変換 / Convert a value typed in current unit to points */
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

// =========================================
// メイン / Main
// =========================================

/* エントリーポイント：ダイアログを表示して整列を実行 / Entry point: show dialog and run alignment */
function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) return;

    var input = showDialog();
    if (input === null) return;

    var corner = input.corner;   // 9点コード：LT TC RT / LM C RM / LB BC RB
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

        try { app.redraw(); } catch (eRedrawMain1) { }

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
        try { app.redraw(); } catch (eRedrawMain2) { }
    }
}

main();

// =========================================
// ダイアログ UI / Dialog UI
// =========================================

/* 設定ダイアログを表示し、入力結果を返す / Show settings dialog and return user input */
function showDialog() {
    var w = new Window("dialog", (lang === "ja") ? LABELS.dialogTitle.ja : LABELS.dialogTitle.en);
    var unitLabel = getCurrentUnitLabel();

    /* プレビュー状態 — ダイアログ閉時に巻き戻し / Preview state (reverted on close) */
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
    } catch (ePrevInit) { }

    function revertPreview(doRedraw) {
        if (doRedraw === undefined) doRedraw = true;
        if (!previewState.applied) return;
        for (var i = 0; i < previewState.items.length; i++) {
            var item = previewState.items[i];
            var dx = previewState.dx[i];
            var dy = previewState.dy[i];
            if (!item) continue;
            try {
                if (dx !== 0 || dy !== 0) item.translate(-dx, -dy);
            } catch (eRev) { }
            previewState.dx[i] = 0;
            previewState.dy[i] = 0;
        }
        previewState.applied = false;
        if (doRedraw) {
            try { app.redraw(); } catch (eRedraw1) { }
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
        try { app.redraw(); } catch (eRedraw2) { }
    }

    function getCornerFromUI() {
        if (rbLT.value) return "LT";
        if (rbTC.value) return "TC";
        if (rbRT.value) return "RT";
        if (rbLM.value) return "LM";
        if (rbC.value) return "C";
        if (rbRM.value) return "RM";
        if (rbLB.value) return "LB";
        if (rbBC.value) return "BC";
        if (rbRB.value) return "RB";
        return "LT";
    }

    function addCornerKeyHandler(dialog) {
        dialog.addEventListener("keydown", function (event) {
            // テキスト入力中はショートカットを無効化 / Skip when typing in edittext
            if (event.target && event.target.type === "edittext") return;
            // 整列先パネルがディムされているときは無効 / Skip when corner panel is disabled
            try { if (!pCorner.enabled) return; } catch (eP) { }

            var k = event.keyName;
            if (!k) return;
            k = ("" + k).toLowerCase();

            var target = null;
            if (k === "q") target = rbLT;
            else if (k === "w") target = rbTC;
            else if (k === "e") target = rbRT;
            else if (k === "a") target = rbLM;
            else if (k === "s") target = rbC;
            else if (k === "d") target = rbRM;
            else if (k === "z") target = rbLB;
            else if (k === "x") target = rbBC;
            else if (k === "c") target = rbRB;

            if (target) {
                selectCornerRb(target);
                event.preventDefault();
                updatePreviewFromUI();
            }
        });
    }

    /* レイアウト：上=整列の基準（全幅）／中=2カラム（整列先・マージン）／下=ボタン
       Layout: top=Base (full width) / middle=2-column / bottom=Buttons */
    w.orientation = "column";
    w.alignChildren = ["fill", "top"];

    /* 整列の基準パネル（全幅・横並び） / Alignment-base panel (top, full-width, horizontal radios) */
    var pTarget = w.add("panel", undefined, _(LABELS.target.ja, LABELS.target.en));
    pTarget.orientation = "row";
    pTarget.alignChildren = ["left", "center"];
    pTarget.margins = [15, 20, 15, 10];

    var rbTargetArtboard = pTarget.add("radiobutton", undefined, _(LABELS.targetArtboard.ja, LABELS.targetArtboard.en));
    var rbTargetSelection = pTarget.add("radiobutton", undefined, _(LABELS.targetSelection.ja, LABELS.targetSelection.en));
    rbTargetSelection.value = true;

    // Helper: check if target is "Selection"
    function isTargetSelection() {
        try { return rbTargetSelection.value === true; } catch (e) { return false; }
    }

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
    pCorner.alignChildren = ["center", "top"];
    pCorner.margins = [15, 20, 15, 10];
    pCorner.spacing = 6;

    var rowTop = pCorner.add("group");
    rowTop.orientation = "row";
    rowTop.alignChildren = ["center", "center"];
    rowTop.spacing = 12;
    var rbLT = rowTop.add("radiobutton", undefined, "");
    var rbTC = rowTop.add("radiobutton", undefined, "");
    var rbRT = rowTop.add("radiobutton", undefined, "");

    var rowMid = pCorner.add("group");
    rowMid.orientation = "row";
    rowMid.alignChildren = ["center", "center"];
    rowMid.spacing = 12;
    var rbLM = rowMid.add("radiobutton", undefined, "");
    var rbC = rowMid.add("radiobutton", undefined, "");
    var rbRM = rowMid.add("radiobutton", undefined, "");

    var rowBot = pCorner.add("group");
    rowBot.orientation = "row";
    rowBot.alignChildren = ["center", "center"];
    rowBot.spacing = 12;
    var rbLB = rowBot.add("radiobutton", undefined, "");
    var rbBC = rowBot.add("radiobutton", undefined, "");
    var rbRB = rowBot.add("radiobutton", undefined, "");

    rbLT.helpTip = _(LABELS.leftTop.ja, LABELS.leftTop.en) + " (q)";
    rbTC.helpTip = _(LABELS.topCenter.ja, LABELS.topCenter.en) + " (w)";
    rbRT.helpTip = _(LABELS.rightTop.ja, LABELS.rightTop.en) + " (e)";
    rbLM.helpTip = _(LABELS.leftMiddle.ja, LABELS.leftMiddle.en) + " (a)";
    rbC.helpTip = _(LABELS.center.ja, LABELS.center.en) + " (s)";
    rbRM.helpTip = _(LABELS.rightMiddle.ja, LABELS.rightMiddle.en) + " (d)";
    rbLB.helpTip = _(LABELS.leftBottom.ja, LABELS.leftBottom.en) + " (z)";
    rbBC.helpTip = _(LABELS.bottomCenter.ja, LABELS.bottomCenter.en) + " (x)";
    rbRB.helpTip = _(LABELS.rightBottom.ja, LABELS.rightBottom.en) + " (c)";

    var allCornerRadios = [rbLT, rbTC, rbRT, rbLM, rbC, rbRM, rbLB, rbBC, rbRB];

    function selectCornerRb(target) {
        for (var i = 0; i < allCornerRadios.length; i++) {
            allCornerRadios[i].value = (allCornerRadios[i] === target);
        }
    }

    selectCornerRb(rbLT);
    addCornerKeyHandler(w);
    // Target shortcut keys: 1 = Artboard, 2 = Selection
    w.addEventListener("keydown", function (event) {
        if (event.target && event.target.type === "edittext") return;
        var k = event.keyName;
        if (!k) return;
        k = ("" + k).toLowerCase();

        if (k === "1") {
            rbTargetArtboard.value = true;
            rbTargetSelection.value = false;
            event.preventDefault();
            updatePreviewFromUI();
        } else if (k === "2") {
            rbTargetSelection.value = true;
            rbTargetArtboard.value = false;
            event.preventDefault();
            updatePreviewFromUI();
        }
    });

    /* コーナーラジオは3行に分かれて自動排他が効かないので、クリック時に明示排他化
       Corners span 3 row groups → enforce mutual exclusion explicitly on click */
    function makeCornerClickHandler(rb) {
        return function () {
            selectCornerRb(rb);
            updatePreviewFromUI();
        };
    }
    for (var i = 0; i < allCornerRadios.length; i++) {
        allCornerRadios[i].onClick = makeCornerClickHandler(allCornerRadios[i]);
    }
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

    function onMarginXEdit() {
        if (cbLink.value) etMarginY.text = etMarginX.text;
        updatePreviewFromUI();
    }
    function onMarginYEdit() {
        if (!cbLink.value) updatePreviewFromUI();
    }
    etMarginX.onChange = etMarginX.onChanging = onMarginXEdit;
    etMarginY.onChange = etMarginY.onChanging = onMarginYEdit;

    // Enter-to-OK for both fields
    function bindEnterToOK(editText) {
        editText.addEventListener("keydown", function (ev) {
            if (ev.keyName === "Enter" || ev.keyName === "Return") {
                try { btnOK.notify(); } catch (eNotify) { }
                ev.preventDefault();
            }
        });
    }
    /* オプション：プレビュー境界 / Option: use preview bounds */
    var gOptions = w.add("group");
    gOptions.orientation = "row";
    gOptions.alignment = ["left", "top"];
    gOptions.margins = [4, 4, 4, 4];

    var cbUsePreviewBounds = gOptions.add("checkbox", undefined, _(LABELS.usePreviewBounds.ja, LABELS.usePreviewBounds.en));
    cbUsePreviewBounds.value = usePreviewBounds;
    cbUsePreviewBounds.onClick = function () {
        usePreviewBounds = cbUsePreviewBounds.value;
        updatePreviewFromUI();
    };

    // --- Buttons (bottom, centered, side-by-side)
    var gButtons = w.add("group");
    gButtons.orientation = "row";
    gButtons.alignChildren = ["center", "center"];
    gButtons.alignment = ["center", "bottom"];

    var btnCancel = gButtons.add("button", undefined, _(LABELS.cancel.ja, LABELS.cancel.en), { name: "cancel" });
    var btnOK = gButtons.add("button", undefined, _(LABELS.ok.ja, LABELS.ok.en), { name: "ok" });

    try { w.defaultElement = btnOK; } catch (eDef) { }

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

    /* モードに応じて整列先パネル／マージン欄のディム状態を同期
       Sync enabled state of corner panel, margin panel, link checkbox */
    function syncDialogEnabledState(targetIsSelection, isCenter) {
        try { pCorner.enabled = !targetIsSelection; } catch (eDimCorner) { }
        var enableMargin = (!isCenter && !targetIsSelection);
        try { pMargin.enabled = enableMargin; } catch (eDim) { }
        try { cbLink.enabled = enableMargin; } catch (eDim2) { }
    }

    /* 入力欄からマージンを pt 単位で取得（中央・選択モード時は 0）
       Read margin from input fields as pt (0 when Center or Selection mode) */
    function readMarginPt(isCenter, targetIsSelection) {
        if (isCenter || targetIsSelection) return { x: 0, y: 0 };
        var mxInput = safeToNumber(etMarginX.text, 0);
        var myInput = safeToNumber(etMarginY.text, 0);
        var sourceY = cbLink.value ? mxInput : myInput;
        var mxPt = toPt(mxInput, unitLabel);
        var myPt = toPt(sourceY, unitLabel);
        return { x: (mxPt === null ? 0 : mxPt), y: (myPt === null ? 0 : myPt) };
    }

    /* UI状態を読み取り、ディム同期とプレビュー反映までを一括実行
       Read UI state, sync dimming, and trigger preview */
    function updatePreviewFromUI() {
        var targetIsSelection = isTargetSelection();
        var corner = getCornerFromUI();
        var isCenter = (corner === "C");

        syncDialogEnabledState(targetIsSelection, isCenter);

        var margin = readMarginPt(isCenter, targetIsSelection);
        applyPreview(corner, margin.x, margin.y, targetIsSelection ? "selection" : "artboard");
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
        } catch (eFocus) { }
    };

    // 初期表示でプレビュー適用（選択がある場合）
    updateLinkStateUI();

    btnOK.onClick = function () {
        closingByOK = true;
        revertPreview(false);
        w.close(1);
    };

    if (w.show() !== 1) return null;

    var corner = getCornerFromUI();
    var isCenter = (corner === "C");
    var targetIsSelection = isTargetSelection();
    var marginPt = readMarginPt(isCenter, targetIsSelection);
    return {
        corner: corner,
        targetMode: targetIsSelection ? "selection" : "artboard",
        // 「アクティブを基準」または「中央」のときは 0、それ以外は { x, y } / 0 when Selection or Center, otherwise {x, y}
        marginPt: (isCenter || targetIsSelection) ? 0 : marginPt
    };
}

/* 文字列を安全に数値化、不正値は fallback / Safe string-to-number with fallback (default: null) */
function safeToNumber(s, fallback) {
    if (fallback === undefined) fallback = null;
    if (s === null || s === undefined) return fallback;
    s = ("" + s).replace(/^\s+|\s+$/g, "");
    if (s === "") return fallback;
    var n = Number(s);
    if (isNaN(n)) return fallback;
    return n;
}

/* 数値入力欄を矢印キーで増減（Shift=10、Option=0.1） / Arrow-key increment for number fields */
function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
    if (allowNegative === undefined) allowNegative = false;

    editText.addEventListener("keydown", function (event) {
        // Up / Down 以外は素通り（数値入力を妨げない） / Pass through non-arrow keys
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = Number(editText.text);
        if (isNaN(value)) value = 0;

        var keyboard = ScriptUI.environment.keyboardState;
        var direction = (event.keyName === "Up") ? 1 : -1;

        if (keyboard.shiftKey) {
            // Shift：10の倍数にスナップ / Snap to multiples of 10
            var step = 10;
            value = (direction > 0)
                ? Math.ceil((value + 1) / step) * step
                : Math.floor((value - 1) / step) * step;
        } else if (keyboard.altKey) {
            // Option：0.1単位 / Increment by 0.1
            value = Math.round((value + direction * 0.1) * 10) / 10;
        } else {
            // 通常：1単位、整数丸め / Default: by 1, integer
            value = Math.round(value + direction);
        }

        if (!allowNegative && value < 0) value = 0;
        editText.text = value;

        if (typeof onValueChanged === "function") {
            try { onValueChanged(); } catch (eCb) { }
        }

        event.preventDefault();
    });
}

// =========================================
// アートボード振り分けと判定 / Artboard mapping & predicates
// =========================================

/* 選択アイテムを中心点が属するアートボードごとに振り分け / Group selection by containing artboard (by center) */
function groupSelectedItemsByArtboard(doc, selection) {
    var result = {};

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (isLockedOrHidden(item)) continue;

        var itemBounds = getItemBounds(item); // [L,T,R,B]
        if (!itemBounds) continue;

        var cx = (itemBounds[0] + itemBounds[2]) / 2;
        var cy = (itemBounds[1] + itemBounds[3]) / 2;

        var abIndex = findArtboardIndexByPoint(doc, cx, cy);
        if (abIndex === -1) continue;

        if (!result[abIndex]) result[abIndex] = [];
        result[abIndex].push(item);
    }

    return result;
}

/* 座標(x,y)を含むアートボードのインデックスを返す / Return index of artboard containing (x, y) */
function findArtboardIndexByPoint(doc, x, y) {
    var n = doc.artboards.length;
    for (var i = 0; i < n; i++) {
        var r = doc.artboards[i].artboardRect; // [L, T, R, B]
        if (x >= r[0] && x <= r[2] && y <= r[1] && y >= r[3]) return i;
    }
    return -1;
}

/* アイテムまたは親階層がロック/非表示か判定 / Check if item or any ancestor is locked/hidden */
function isLockedOrHidden(item) {
    try { if (item.locked) return true; } catch (e) { }
    try { if (item.hidden) return true; } catch (e2) { }

    var p = item.parent;
    while (p) {
        try { if (p.locked) return true; } catch (e3) { }
        try { if (p.hidden) return true; } catch (e4) { }
        if (p.typename === "Document") break;
        p = p.parent;
    }
    return false;
}

/* 複数アイテムを包含する最小バウンディングを返す / Return union bounds of multiple items */
function getUnionBounds(items) {
    if (!items || items.length === 0) return null;

    var l = null, t = null, r = null, b = null;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        var itemBounds = getItemBounds(item);
        if (!itemBounds) continue;

        if (l === null) {
            l = itemBounds[0]; t = itemBounds[1]; r = itemBounds[2]; b = itemBounds[3];
        } else {
            if (itemBounds[0] < l) l = itemBounds[0];
            if (itemBounds[1] > t) t = itemBounds[1];
            if (itemBounds[2] > r) r = itemBounds[2];
            if (itemBounds[3] < b) b = itemBounds[3];
        }
    }

    if (l === null) return null;
    return [l, t, r, b];
}

// =========================================
// 整列処理 / Alignment
// =========================================

/* 整列先アンカーのアートボード上の座標を計算 / Compute target anchor point on the artboard */
function computeCornerTargetPoint(artboardRect, corner, marginX, marginY) {
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    if (corner === "LT") return [abL + marginX, abT - marginY];
    if (corner === "TC") return [abCx, abT - marginY];
    if (corner === "RT") return [abR - marginX, abT - marginY];
    if (corner === "LM") return [abL + marginX, abCy];
    if (corner === "C") return [abCx, abCy];
    if (corner === "RM") return [abR - marginX, abCy];
    if (corner === "LB") return [abL + marginX, abB + marginY];
    if (corner === "BC") return [abCx, abB + marginY];
    return [abR - marginX, abB + marginY]; // RB
}

/* 各アイテムをアートボード基準のアンカーへ整列（マージン付き） / Align each item to artboard anchor with margin */
function alignItemsToCornerWithMargin(items, artboardRect, corner, margin) {
    // margin: pt (number or {x, y})
    var marginX = 0, marginY = 0;
    if (typeof margin === "object" && margin !== null) {
        marginX = margin.x || 0;
        marginY = margin.y || 0;
    } else {
        marginX = marginY = margin || 0;
    }
    var target = computeCornerTargetPoint(artboardRect, corner, marginX, marginY);
    var targetX = target[0], targetY = target[1];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var itemBounds = getItemBounds(item); // [L,T,R,B]
        if (!itemBounds) continue;
        var anchor = getBoundsAnchorPoint(itemBounds, corner);
        var dx = targetX - anchor[0];
        var dy = targetY - anchor[1];
        try { item.translate(dx, dy); } catch (e2) { }
    }
}

/* プレビュー用：移動量を記録しながら整列（後で巻き戻し可能） / Preview-mode alignment that records deltas for revert */
function alignItemsToCornerWithMarginPreviewXY(items, artboardRect, corner, marginX, marginY, previewState) {
    // marginX, marginY: pt
    var target = computeCornerTargetPoint(artboardRect, corner, marginX, marginY);
    var targetX = target[0], targetY = target[1];

    // Record per-item delta so we can revert on close
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;
        var itemBounds = getItemBounds(item); // [L,T,R,B]
        if (!itemBounds) continue;
        var anchor = getBoundsAnchorPoint(itemBounds, corner);
        var dx = targetX - anchor[0];
        var dy = targetY - anchor[1];
        // Find index in previewState.items
        var idx = -1;
        for (var j = 0; j < previewState.items.length; j++) {
            if (previewState.items[j] === item) { idx = j; break; }
        }
        try { item.translate(dx, dy); } catch (e2) { }
        if (idx >= 0) {
            previewState.dx[idx] += dx;
            previewState.dy[idx] += dy;
        }
    }
}

/* アートボード矩形上の9点アンカー座標を返す / Return one of the 9 anchor points on the artboard rect */
function getArtboardAnchorPoint(artboardRect, corner) {
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    if (corner === "C") return [abCx, abCy];
    if (corner === "LT") return [abL, abT];
    if (corner === "TC") return [abCx, abT];
    if (corner === "RT") return [abR, abT];
    if (corner === "LM") return [abL, abCy];
    if (corner === "RM") return [abR, abCy];
    if (corner === "LB") return [abL, abB];
    if (corner === "BC") return [abCx, abB];
    return [abR, abB]; // RB
}

/* バウンディングボックス上の9点アンカー座標を返す / Return one of the 9 anchor points on a bounds rect */
function getBoundsAnchorPoint(bounds, corner) {
    // bounds: [L,T,R,B]
    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3];
    var cx = (l + r) / 2;
    var cy = (t + b) / 2;

    if (corner === "C") return [cx, cy];
    if (corner === "LT") return [l, t];
    if (corner === "TC") return [cx, t];
    if (corner === "RT") return [r, t];
    if (corner === "LM") return [l, cy];
    if (corner === "RM") return [r, cy];
    if (corner === "LB") return [l, b];
    if (corner === "BC") return [cx, b];
    return [r, b]; // RB
}

/* テンプレ群とアートボードアンカーの相対オフセットを計算 / Compute relative offset of template group from artboard anchor */
function computeGroupOffsetFromArtboard(items, artboardRect, corner) {
    // Returns {x,y} offset = (groupAnchor - artboardAnchor)
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return null;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);

    return { x: (groupAnchor[0] - artboardAnchor[0]), y: (groupAnchor[1] - artboardAnchor[1]) };
}

/* アイテム群をオフセット付きでアートボードアンカーに整列 / Move group so its anchor equals (artboardAnchor + offset) */
function alignGroupToArtboardWithOffset(items, artboardRect, corner, offset) {
    // Moves all items together so that their groupAnchor equals (artboardAnchor + offset)
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);
    var dx = artboardAnchor[0] + (offset.x || 0) - groupAnchor[0];
    var dy = artboardAnchor[1] + (offset.y || 0) - groupAnchor[1];

    for (var i = 0; i < items.length; i++) {
        try { items[i].translate(dx, dy); } catch (e) { }
    }
}

/* プレビュー用：アイテム群整列＋移動量を記録 / Preview-mode group alignment that records deltas for revert */
function alignGroupToArtboardWithOffsetPreview(items, artboardRect, corner, offset, previewState) {
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);
    var dx = artboardAnchor[0] + (offset.x || 0) - groupAnchor[0];
    var dy = artboardAnchor[1] + (offset.y || 0) - groupAnchor[1];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        var idx = -1;
        for (var j = 0; j < previewState.items.length; j++) {
            if (previewState.items[j] === item) { idx = j; break; }
        }

        try { item.translate(dx, dy); } catch (e2) { }
        if (idx >= 0) {
            previewState.dx[idx] += dx;
            previewState.dy[idx] += dy;
        }
    }
}