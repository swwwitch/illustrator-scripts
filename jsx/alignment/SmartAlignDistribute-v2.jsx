#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartAlignDistribute.jsx

### オリジナルアイデア

John Wundes 
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

*/

/* バージョン変数を追加 / Script version variable */
var SCRIPT_VERSION = "v1.0.2";

/* ダイアログ外観変数 / Dialog appearance variable */
var dialogOpacity = 0.97;

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

/* LABELS 定義（ダイアログUIの順序に合わせて並べ替え）/ Label definitions (ordered to match dialog UI) */
var LABELS = {
    dialogTitle: {
        ja: "整列と分布" + SCRIPT_VERSION,
        en: "Align & Distribute " + SCRIPT_VERSION
    },
    spacingLabel: {
        ja: "間隔",
        en: "Spacing"
    },
    directionPanel: { ja: "方向", en: "Direction" },
    dirAuto: { ja: "自動", en: "Auto" },
    dirVertical: { ja: "縦", en: "Vertical" },
    dirHorizontal: { ja: "横", en: "Horizontal" },

    alignH: { ja: "揃え（左右）", en: "Align (H)" },
    alignV: { ja: "揃え（上下）", en: "Align (V)" },

    vAlignTop: { ja: "上", en: "Top" },
    vAlignMiddle: { ja: "中央", en: "Middle" },
    vAlignBottom: { ja: "下", en: "Bottom" },
    vAlignNone: { ja: "なし", en: "None" },
    useBounds: {
        ja: "プレビュー境界を使用",
        en: "Use preview bounds"
    },
    random: {
        ja: "ランダム",
        en: "Random"
    },
    hAlignLeft: {
        ja: "左",
        en: "Left"
    },
    hAlignCenter: {
        ja: "中央",
        en: "Center"
    },
    hAlignRight: {
        ja: "右",
        en: "Right"
    },
    hAlignNone: {
        ja: "なし",
        en: "None"
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

/* 位置復元式プレビュー管理 / Position-restore preview manager (no Undo dependence) */
function PositionPreviewManager(items) {
    this.items = (items && items.length) ? items.slice() : [];
    this.positions = [];

    // Snapshot initial positions
    for (var i = 0; i < this.items.length; i++) {
        var it = this.items[i];
        try {
            this.positions.push([it.left, it.top]);
        } catch (e) {
            this.positions.push([0, 0]);
        }
    }

    // Restore to snapshot positions
    this.restore = function () {
        for (var i = 0; i < this.items.length; i++) {
            var it = this.items[i];
            if (!it) continue;
            var p = this.positions[i];
            if (!p) continue;
            try {
                it.left = p[0];
                it.top = p[1];
            } catch (e) { }
        }
        try { app.redraw(); } catch (e) { }
    };

    // Confirm: keep current positions (no-op)
    this.confirm = function () {
        // no-op
    };
}

/* 単位コードとラベルのマップ / Map of unit codes to labels */
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

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/* 単位コードからポイント換算係数を取得 / Get point factor from unit code */
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

function addAlignKeyHandler(target, getDir, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, onUpdate) {
    target.addEventListener("keydown", function (event) {
        var k = event.keyName;
        var dir = (typeof getDir === "function") ? getDir() : "vertical";

        function applyAndUpdate() {
            event.preventDefault();
            if (typeof onUpdate === "function") onUpdate();
        }

        // "なし" (both directions)
        if (k === "N") {
            if (dir === "horizontal") rbVNone.value = true;
            else rbHNone.value = true;
            applyAndUpdate();
            return;
        }

        // Vertical stacking: Left/Center/Right by L/C/R
        if (dir !== "horizontal") {
            if (k === "L") { rbHLeft.value = true; applyAndUpdate(); return; }
            if (k === "C") { rbHCenter.value = true; applyAndUpdate(); return; }
            if (k === "R") { rbHRight.value = true; applyAndUpdate(); return; }
            return;
        }

        // Horizontal stacking: Top/Middle/Bottom by T/M/B
        if (k === "T") { rbVTop.value = true; applyAndUpdate(); return; }
        if (k === "M") { rbVMiddle.value = true; applyAndUpdate(); return; }
        if (k === "B") { rbVBottom.value = true; applyAndUpdate(); return; }
    });
}

/* ダイアログの不透明度を設定するヘルパー / Helper to set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
}

/**
 * Collapse or expand a group so it doesn't reserve space when hidden.
 * On first use, stores original minimumSize, maximumSize, and preferredSize.
 */
function setGroupCollapsed(g, collapsed) {
    if (!g) return;
    // store original sizes once
    if (g.__origMinSize == null) {
        g.__origMinSize = [g.minimumSize.width, g.minimumSize.height];
        g.__origMaxSize = [g.maximumSize.width, g.maximumSize.height];
        g.__origPrefSize = [g.preferredSize.width, g.preferredSize.height];
    }

    if (collapsed) {
        g.visible = false;
        g.enabled = false;
        g.minimumSize = [0, 0];
        g.preferredSize = [0, 0];
        // maximumSize は触らない（クラッシュ回避）
    } else {
        g.visible = true;
        g.enabled = true;
        g.minimumSize = g.__origMinSize;
        g.maximumSize = g.__origMaxSize;
        g.preferredSize = g.__origPrefSize;
    }
}


/* ダイアログ位置（セッション内で記憶） / Dialog position (remember within session) */
var DLG_POS_MEM_KEY = "__SmartAlignAndTileTateSimple_DlgPos__";

function loadDialogPosition() {
    try {
        var p = $.global[DLG_POS_MEM_KEY];
        if (p && p.length === 2) return [p[0], p[1]];
    } catch (e) { }
    return null;
}

function saveDialogPosition(loc) {
    try {
        if (!loc || loc.length !== 2) return;
        $.global[DLG_POS_MEM_KEY] = [Math.round(loc[0]), Math.round(loc[1])];
    } catch (e) { }
}

/* アイテムの境界を取得（プレビュー境界ONならvisible、OFFならgeometric） / Get item bounds (visible when preview ON, geometric when OFF) */
function getItemBounds(item, usePreviewBounds) {
    return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
}

function getSelectionSpan(items, usePreviewBounds) {
    var minL = null, maxR = null, maxT = null, minB = null;
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var b = getItemBounds(it, usePreviewBounds);
        var l = b[0], t = b[1], r = b[2], bo = b[3];
        if (minL === null || l < minL) minL = l;
        if (maxR === null || r > maxR) maxR = r;
        if (maxT === null || t > maxT) maxT = t;
        if (minB === null || bo < minB) minB = bo;
    }
    if (minL === null) return { spanX: 0, spanY: 0 };
    return { spanX: (maxR - minL), spanY: (maxT - minB) };
}

function detectDirection(items) {
    // geometricBoundsで安定判定
    var s = getSelectionSpan(items, false);
    return (s.spanX >= s.spanY) ? "horizontal" : "vertical";
}

/* ダイアログ表示関数 / Show dialog with language support */
function showArrangeDialog() {
    var lang = getCurrentLang();
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    /* ダイアログの不透明度と位置を設定 / Set dialog opacity and position */
    setDialogOpacity(dlg, dialogOpacity);
    var lastPos = loadDialogPosition();
    if (lastPos) dlg.location = lastPos;

    // Preserve current preference so cancel can restore it
    var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

    // Snapshot selection
    var originalSelection = activeDocument.selection.slice();

    // Position-restore preview manager (no Undo)
    var previewMgr = new PositionPreviewManager(originalSelection);

    var detectedDir = detectDirection(originalSelection);

    // Direction UI
    var dirPanel = dlg.add("panel", undefined, LABELS.directionPanel[lang]);
    dirPanel.orientation = "row";
    dirPanel.alignChildren = ["center", "center"];
    dirPanel.margins = [15, 20, 15, 10];

    var rbDirAuto = dirPanel.add("radiobutton", undefined, LABELS.dirAuto[lang]);
    var rbDirV = dirPanel.add("radiobutton", undefined, LABELS.dirVertical[lang]);
    var rbDirH = dirPanel.add("radiobutton", undefined, LABELS.dirHorizontal[lang]);
    rbDirAuto.value = true;


    function getEffectiveDirection() {
        if (rbDirV.value) return "vertical";
        if (rbDirH.value) return "horizontal";
        return detectedDir;
    }

    /* 間隔 / Spacing */
    var unit = getCurrentUnitLabel();
    var spacingPanel = dlg.add("panel", undefined, LABELS.spacingLabel[lang]);
    spacingPanel.orientation = "column";
    spacingPanel.alignChildren = ["center", "center"];
    spacingPanel.margins = [15, 20, 15, 10];

    // 垂直方向（間隔） / Vertical spacing
    var hMarginGroup = spacingPanel.add("group");
    hMarginGroup.orientation = "row";
    hMarginGroup.alignChildren = ["left", "center"];

    var hMarginInput = hMarginGroup.add("edittext", undefined, "0");
    hMarginInput.characters = 3;
    var hMarginUnitLabel = hMarginGroup.add("statictext", undefined, unit);

    changeValueByArrowKey(hMarginInput, true, function () {
        requestPreviewUpdate();
    });

    /* 揃えパネル / Align panel */
    var alignPanel = dlg.add("panel", undefined, "");
    alignPanel.orientation = "column";
    alignPanel.alignChildren = ["left", "center"];
    alignPanel.margins = [15, 20, 15, 10];

    /* 左右方向揃えグループ（基本はこちらを使用） / Horizontal align group (primary) */
    var hAlignGroup = alignPanel.add("group");
    hAlignGroup.orientation = "row";
    hAlignGroup.alignChildren = ["left", "center"];
    var rbHNone = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignNone[lang]);
    var rbHLeft = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignLeft[lang]);
    var rbHCenter = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignCenter[lang]);
    var rbHRight = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignRight[lang]);
    rbHNone.value = true; /* デフォルトを「なし」に / Default to None */

    // 上下揃え（横並び時） / Vertical align for horizontal layout
    var vAlignGroup = alignPanel.add("group");
    vAlignGroup.orientation = "row";
    vAlignGroup.alignChildren = ["left", "center"];

    var rbVNone = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignNone[lang]);
    var rbVTop = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignTop[lang]);
    var rbVMiddle = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignMiddle[lang]);
    var rbVBottom = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignBottom[lang]);
    rbVNone.value = true;

    rbVNone.onClick = function () { requestPreviewUpdate(); };
    rbVTop.onClick = function () { requestPreviewUpdate(); };
    rbVMiddle.onClick = function () { requestPreviewUpdate(); };
    rbVBottom.onClick = function () { requestPreviewUpdate(); };

    function syncAlignUI() {
        var d = getEffectiveDirection();

        // 2段とも表示し、使わない方はディムにする
        hAlignGroup.visible = true;
        vAlignGroup.visible = true;

        if (d === "horizontal") {
            alignPanel.text = LABELS.alignV[lang];
            hAlignGroup.enabled = false; // ディム
            vAlignGroup.enabled = true;
        } else {
            alignPanel.text = LABELS.alignH[lang];
            hAlignGroup.enabled = true;
            vAlignGroup.enabled = false; // ディム
        }

        try { dlg.layout.layout(true); } catch (e) { }
    }

    rbHLeft.onClick = function () { requestPreviewUpdate(); };
    rbHCenter.onClick = function () { requestPreviewUpdate(); };
    rbHRight.onClick = function () { requestPreviewUpdate(); };
    rbHNone.onClick = function () { requestPreviewUpdate(); };

    /* オプション（プレビュー境界/ランダム） / Options (bounds/random) */
    var optGroup = dlg.add("group");
    optGroup.orientation = "column";
    optGroup.alignChildren = ["left", "center"];
    optGroup.alignment = ["fill", "top"];
    optGroup.margins = [15, 5, 15, 5];

    var boundsCheckbox = optGroup.add("checkbox", undefined, LABELS.useBounds[lang]);
    boundsCheckbox.value = true;
    boundsCheckbox.onClick = function () {
        requestPreviewUpdate();
    };

    var randomCheckbox = optGroup.add("checkbox", undefined, LABELS.random[lang]);
    randomCheckbox.value = false;

    /* ボタン配置グループ / Button group */
    var buttonGroup2 = dlg.add("group");
    buttonGroup2.alignment = "center";
    buttonGroup2.alignChildren = ["center", "center"];
    var cancelButton = buttonGroup2.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var okButton = buttonGroup2.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });


    // Random preview cache: keep the same shuffle order so OK matches preview
    var randomOrderCache = null; // array of indices

    function applyLayoutToSelection() {
        // Guard
        if (!originalSelection || originalSelection.length === 0) return;

        // 垂直方向のみ / Vertical only
        var vMarginValue = parseFloat(hMarginInput.text);
        if (isNaN(vMarginValue)) vMarginValue = 0;

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var ptFactor = getPtFactorFromUnitCode(unitCode);
        var vMarginPt = vMarginValue * ptFactor;

        // Sort items (or shuffle)
        var sortedItems;
        var baseLeft = null;
        var baseTop = null;

        if (randomCheckbox.value) {
            // Build or reuse a stable shuffle order so preview == final on OK
            if (!randomOrderCache || randomOrderCache.length !== originalSelection.length) {
                randomOrderCache = [];
                for (var ri = 0; ri < originalSelection.length; ri++) randomOrderCache.push(ri);
                for (var i = randomOrderCache.length - 1; i > 0; i--) {
                    var j = Math.floor(Math.random() * (i + 1));
                    var tmp = randomOrderCache[i];
                    randomOrderCache[i] = randomOrderCache[j];
                    randomOrderCache[j] = tmp;
                }
            }
            sortedItems = [];
            for (var si = 0; si < randomOrderCache.length; si++) {
                sortedItems.push(originalSelection[randomOrderCache[si]]);
            }

            // Preserve a stable base position (top-left) from the current (rolled-back) state
            for (var k = 0; k < originalSelection.length; k++) {
                var it = originalSelection[k];
                if (!it) continue;
                if (baseLeft === null || it.left < baseLeft) baseLeft = it.left;
                if (baseTop === null || it.top > baseTop) baseTop = it.top;
            }
            if (baseLeft === null) baseLeft = sortedItems[0].left;
            if (baseTop === null) baseTop = sortedItems[0].top;

        } else {
            // Not random: clear cache so next time random starts fresh
            randomOrderCache = null;

            // Not random: sort by direction
            var d0 = getEffectiveDirection();
            sortedItems = (d0 === "horizontal") ? sortByX(originalSelection) : sortByY(originalSelection);
        }

        // Use selected bounds type (visible/geometric)
        var startBounds = getItemBounds(sortedItems[0], boundsCheckbox.value);
        var startLeftBound = startBounds[0];
        var startTopBound = startBounds[1];
        var startRightBound = startBounds[2];
        var startBottomBound = startBounds[3];
        var startX = startLeftBound;
        var startY = startTopBound;

        // Reference width (max) for alignment
        var refWidth = 0;
        for (var m = 0; m < sortedItems.length; m++) {
            var vb0 = getItemBounds(sortedItems[m], boundsCheckbox.value);
            var w0 = vb0[2] - vb0[0];
            if (w0 > refWidth) refWidth = w0;
        }

        var dir = getEffectiveDirection();

        if (dir === "horizontal") {
            // 横並び（左→右）
            var refHeight = 0;
            for (var mh = 0; mh < sortedItems.length; mh++) {
                var bb = getItemBounds(sortedItems[mh], boundsCheckbox.value);
                var hh = bb[1] - bb[3];
                if (hh > refHeight) refHeight = hh;
            }

            var currentX = startX;
            for (var idx = 0; idx < sortedItems.length; idx++) {
                var item = sortedItems[idx];
                if (!item) continue;

                var vb = getItemBounds(item, boundsCheckbox.value);
                var itemWidth = vb[2] - vb[0];

                // Xは左基準で配置
                item.left = item.left + (currentX - vb[0]);

                // 上下揃え（なし/上/中央/下）
                var applyVAlign = !rbVNone.value;
                if (applyVAlign) {
                    var cellTop = startY;
                    var cellBottom = startY - refHeight;
                    var cellCenterY = (cellTop + cellBottom) / 2;

                    var currentTop = vb[1];
                    var currentBottom = vb[3];
                    var currentCenterY = (vb[1] + vb[3]) / 2;

                    var deltaY = 0;
                    if (rbVMiddle.value) deltaY = cellCenterY - currentCenterY;
                    else if (rbVBottom.value) deltaY = cellBottom - currentBottom;
                    else deltaY = cellTop - currentTop; // Top

                    item.top = item.top + deltaY;
                }

                currentX += itemWidth + vMarginPt; // 間隔をX方向へ
            }

        } else {
            // 縦並び（上→下）
            var currentY = startY;
            for (var idx2 = 0; idx2 < sortedItems.length; idx2++) {
                var item2 = sortedItems[idx2];
                if (!item2) continue;

                var vb2 = getItemBounds(item2, boundsCheckbox.value);
                var itemHeight2 = vb2[1] - vb2[3];

                // 左右揃え（なし/左/中央/右）
                var applyHAlign = !rbHNone.value;
                if (applyHAlign) {
                    var cellLeft = startX;
                    var cellRight = cellLeft + refWidth;
                    var cellCenterX = (cellLeft + cellRight) / 2;

                    var currentLeftX2 = vb2[0];
                    var currentRightX2 = vb2[2];
                    var currentCenterX2 = (vb2[0] + vb2[2]) / 2;

                    var hAlign = "left";
                    if (rbHCenter.value) hAlign = "center";
                    else if (rbHRight.value) hAlign = "right";

                    var deltaX2 = 0;
                    if (hAlign === "center") deltaX2 = cellCenterX - currentCenterX2;
                    else if (hAlign === "right") deltaX2 = cellRight - currentRightX2;
                    else deltaX2 = cellLeft - currentLeftX2;

                    item2.left = item2.left + deltaX2;
                }

                // YはTop揃えで積む
                item2.top = item2.top + (currentY - vb2[1]);

                if (idx2 < sortedItems.length - 1) {
                    currentY -= itemHeight2 + vMarginPt; // 間隔をY方向へ
                }
            }
        }

        // (Old vertical stacking loop removed; handled by direction-aware branches above)

        // Random base correction
        if (randomCheckbox.value && sortedItems.length > 0) {
            var offsetX = baseLeft - sortedItems[0].left;
            var offsetY = baseTop - sortedItems[0].top;
            for (var t = 0; t < sortedItems.length; t++) {
                if (!sortedItems[t]) continue;
                sortedItems[t].left += offsetX;
                sortedItems[t].top += offsetY;
            }
        }
    }

    /* プレビュー更新処理（Undoを汚さない） / Update preview without polluting Undo history */
    var __isPreviewUpdating = false;
    var __lastPreviewAt = 0;
    var __previewMinIntervalMs = 80;

    function updatePreview() {
        if (__isPreviewUpdating) return;

        var now = new Date().getTime();
        if (now - __lastPreviewAt < __previewMinIntervalMs) return;
        __lastPreviewAt = now;

        __isPreviewUpdating = true;
        try {
            // Restore to snapshot, then apply preview
            previewMgr.restore();
            // Apply current preference for bounds calculation (not undoable)
            try {
                app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value);
            } catch (e) { }
            applyLayoutToSelection();
            try { app.redraw(); } catch (e) { }
        } finally {
            __isPreviewUpdating = false;
        }
    }

    // Defer preview updates to avoid running undo/layout inside ScriptUI event handlers (crash mitigation)
    var __previewTaskId = 0;
    $.global.__SAT_updatePreview = updatePreview;

    function requestPreviewUpdate() {
        try {
            if (__previewTaskId) app.cancelTask(__previewTaskId);
        } catch (e) { }
        $.global.__SAT_updatePreview = updatePreview;
        try {
            __previewTaskId = app.scheduleTask('$.global.__SAT_updatePreview && $.global.__SAT_updatePreview();', 60, false);
        } catch (e) {
            // Fallback: run directly
            updatePreview();
        }
    }

    // N/L/C/R/T/M/B で揃えラジオを切替（dlg + 入力欄の両方に付与して取りこぼし防止）
    addAlignKeyHandler(dlg, getEffectiveDirection, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, requestPreviewUpdate);
    addAlignKeyHandler(hMarginInput, getEffectiveDirection, rbHNone, rbHLeft, rbHCenter, rbHRight, rbVNone, rbVTop, rbVMiddle, rbVBottom, requestPreviewUpdate);

    function onDirChanged() {
        syncAlignUI();
        requestPreviewUpdate();
    }
    rbDirAuto.onClick = onDirChanged;
    rbDirV.onClick = onDirChanged;
    rbDirH.onClick = onDirChanged;

    // 初期同期
    syncAlignUI();

    requestPreviewUpdate();
    randomCheckbox.onClick = function () {
        randomOrderCache = null;
        requestPreviewUpdate();
    };

    /* ダイアログを開く前に 垂直方向（1段目）をアクティブにする / Activate vertical spacing (1st field) before showing dialog */
    hMarginInput.active = true;
    var dlgResult = dlg.show();
    try { if (__previewTaskId) app.cancelTask(__previewTaskId); } catch (e) { }
    saveDialogPosition(dlg.location);
    if (dlgResult !== 1) {
        // Cancel: restore positions and restore preference
        previewMgr.restore();
        app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
        app.redraw();
        return null;
    }

    // 垂直方向のみ / Vertical only
    var vMarginValue = parseFloat(hMarginInput.text);
    if (isNaN(vMarginValue)) vMarginValue = 0;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptFactor = getPtFactorFromUnitCode(unitCode);
    var spacingPt = vMarginValue * ptFactor;
    var hMarginPt = 0; // backward compatibility
    var vMarginPt = spacingPt; // backward compatibility

    var hAlign = "left";
    if (rbHNone.value) hAlign = "none";
    else if (rbHCenter.value) hAlign = "center";
    else if (rbHRight.value) hAlign = "right";

    // Confirm: keep the current preview result (positions already applied)
    try { app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value); } catch (e) { }
    previewMgr.confirm();
    try { app.redraw(); } catch (e) { }

    // Return values are kept for compatibility (main() currently just exits)
    return {
        mode: getEffectiveDirection(),
        hMargin: hMarginPt,
        vMargin: vMarginPt,
        spacing: spacingPt,
        random: randomCheckbox.value,
        grid: false,
        hAlign: hAlign
    };
}

/* 編集テキストで上下キーによる数値変更を有効化 / Enable up/down arrow key increment/decrement on edittext inputs */
function changeValueByArrowKey(editText, allowNegative, onUpdate) {
    editText.addEventListener("keydown", function (event) {
        if (editText.text.length === 0) return;
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            var isUp = event.keyName == "Up";
            var delta = 1;

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / Snap to multiples of 10 */
                value = Math.floor(value / 10) * 10;
                delta = 10;
            }

            value += isUp ? delta : -delta;

            /* 負数許可されない場合は0未満を禁止 / Prevent negative if not allowed */
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            if (typeof onUpdate === "function") {
                onUpdate();
            }
        }
    });
}

/* Y座標でソート（上から下、同じなら左から右） / Sort items by Y coordinate (top to bottom, then left) */
function sortByY(items) {
    var copiedItems = items.slice();
    copiedItems.sort(function (a, b) {
        if (a.top !== b.top) return b.top - a.top; // higher top first
        return a.left - b.left;
    });
    return copiedItems;
}

function sortByX(items) {
    var copiedItems = items.slice();
    copiedItems.sort(function (a, b) {
        if (a.left !== b.left) return a.left - b.left;
        return b.top - a.top;
    });
    return copiedItems;
}

/* メイン処理 / Main function */
function main() {
    try {
        var lang = getCurrentLang();
        var selectedItems = activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(
                (lang === "ja"
                    ? "オブジェクトを選択してください。"
                    : "Please select objects."
                )
            );
            return;
        }

        var arrangeOptions = showArrangeDialog();
        if (!arrangeOptions) return;

        /* プレビュー結果をそのまま使用 / Use preview result as final */
        return;
    } catch (e) {
        var lang = getCurrentLang();
        alert(
            (lang === "ja"
                ? "エラーが発生しました: " + e.message
                : "An error has occurred: " + e.message
            )
        );
    }
}

main();