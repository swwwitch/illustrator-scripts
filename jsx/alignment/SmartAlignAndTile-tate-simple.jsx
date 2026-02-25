#targetengine "SmartAlignAndTileEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AlignAndDistribute-Tate-simple.jsx

### オリジナルアイデア

John Wundes 
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

*/

/* バージョン変数を追加 / Script version variable */
var SCRIPT_VERSION = "v1.0.1";

/* ダイアログ外観変数 / Dialog appearance variable */
var dialogOpacity = 0.97;

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

/* LABELS 定義（ダイアログUIの順序に合わせて並べ替え）/ Label definitions (ordered to match dialog UI) */
var LABELS = {
    dialogTitle: {
        ja: "整列と分布（縦）" + SCRIPT_VERSION,
        en: "Align & Distribute (Vertical) " + SCRIPT_VERSION
    },
    hMargin: {
        ja: "垂直方向:",
        en: "V:"
    },
    spacingLabel: {
        ja: "間隔",
        en: "Spacing"
    },
    useBounds: {
        ja: "プレビュー境界を使用",
        en: "Use preview bounds"
    },
    random: {
        ja: "ランダム",
        en: "Random"
    },
    alignPanel: {
        ja: "揃え",
        en: "Align"
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

/* 汎用 Undo/Preview 管理クラス / Generic Undo-safe preview manager */
function PreviewManager() {
    this.undoDepth = 0; // プレビュー中に実行されたアクションの回数 / Number of preview actions executed

    /**
     * 変更操作を実行し、履歴としてカウントする / Execute a change and count it as a preview step
     * @param {Function} func - 実行したい処理（無名関数で渡す） / The action to execute
     */
    this.addStep = function (func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            alert("Preview Error: " + e);
        }
    };

    /**
     * プレビューのために行った変更を全て取り消す（キャンセル時など） / Roll back all preview changes
     */
    this.rollback = function () {
        while (this.undoDepth > 0) {
            app.undo();
            this.undoDepth--;
        }
        app.redraw();
    };

    /**
     * 現在の状態を確定する（OK時） / Confirm current state
     * @param {Function} [finalAction] - (任意) 全てUndoした後に実行する「本番」の処理 / Optional final action
     */
    this.confirm = function (finalAction) {
        if (finalAction) {
            this.rollback();
            finalAction();
            this.undoDepth = 0;
        } else {
            this.undoDepth = 0;
        }
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

function addAlignKeyHandler(target, rbNone, rbLeft, rbCenter, rbRight, onUpdate) {
    target.addEventListener("keydown", function (event) {
        // ScriptUI: event.keyName は "N" などで来ることが多い
        var k = event.keyName;

        if (k === "N") {
            rbNone.value = true;
            event.preventDefault();
            if (typeof onUpdate === "function") onUpdate();
        } else if (k === "L") {
            rbLeft.value = true;
            event.preventDefault();
            if (typeof onUpdate === "function") onUpdate();
        } else if (k === "C") {
            rbCenter.value = true;
            event.preventDefault();
            if (typeof onUpdate === "function") onUpdate();
        } else if (k === "R") {
            rbRight.value = true;
            event.preventDefault();
            if (typeof onUpdate === "function") onUpdate();
        }
    });
}

/* ダイアログの不透明度を設定するヘルパー / Helper to set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
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

    // Undo-safe preview manager
    var previewMgr = new PreviewManager();

    // Preserve current preference so cancel can restore it
    var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

    /* 間隔 / Spacing */
    var unit = getCurrentUnitLabel();
    var spacingPanel = dlg.add("group");
    spacingPanel.orientation = "column";
    spacingPanel.alignChildren = ["center", "center"];
    spacingPanel.margins = [15, 5, 15, 5];

    // 垂直方向（間隔） / Vertical spacing
    var hMarginGroup = spacingPanel.add("group");
    hMarginGroup.orientation = "row";
    hMarginGroup.alignChildren = ["left", "center"];

    var hMarginLabel = hMarginGroup.add("statictext", undefined, LABELS.spacingLabel[lang]);
    var hMarginInput = hMarginGroup.add("edittext", undefined, "0");
    hMarginInput.characters = 3;
    var hMarginUnitLabel = hMarginGroup.add("statictext", undefined, unit);

    changeValueByArrowKey(hMarginInput, true, function () {
        updatePreview();
    });

    /* 揃えパネル / Align panel */
    var alignPanel = dlg.add("panel", undefined, LABELS.alignPanel[lang]);
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

    rbHLeft.onClick = function () { updatePreview(); };
    rbHCenter.onClick = function () { updatePreview(); };
    rbHRight.onClick = function () { updatePreview(); };
    rbHNone.onClick = function () { updatePreview(); };

    /* オプション（プレビュー境界/ランダム） / Options (bounds/random) */
    var optGroup = dlg.add("group");
    optGroup.orientation = "column";
    optGroup.alignChildren = ["left", "center"];
    optGroup.alignment = ["fill", "top"];
    optGroup.margins = [15, 5, 15, 5];

    var boundsCheckbox = optGroup.add("checkbox", undefined, LABELS.useBounds[lang]);
    boundsCheckbox.value = true;
    boundsCheckbox.onClick = function () {
        updatePreview();
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

    // Keep a snapshot of the selection reference array (used for preview/final)
    var originalSelection = activeDocument.selection.slice();

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
        var hMarginPt = 0;
        var vMarginPt = vMarginValue * ptFactor;

        // Always single column (cols UI removed)
        var cols = 1;
        var itemsPerCol = originalSelection.length;

        // Horizontal align (primary)
        var applyHAlign = !rbHNone.value;
        var hAlign = "left";
        if (rbHCenter.value) hAlign = "center";
        else if (rbHRight.value) hAlign = "right";

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

            // Vertical-first: sort by Y (top descending), then by X (left ascending)
            sortedItems = sortByY(originalSelection);
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

        var index = 0;
        for (var c = 0; c < cols; c++) {
            var currentY = startY;
            for (var r = 0; r < itemsPerCol && index < sortedItems.length; r++, index++) {
                var item = sortedItems[index];
                if (!item) continue;

                var vb = getItemBounds(item, boundsCheckbox.value);
                var itemWidth = vb[2] - vb[0];
                var itemHeight = vb[1] - vb[3];

                // Cell X range (use fixed column width so Left/Center/Right works)
                var cellLeft = startX + c * (refWidth + hMarginPt);
                var cellRight = cellLeft + refWidth;
                var cellCenterX = (cellLeft + cellRight) / 2;

                // Current anchors (X)
                var currentLeftX = vb[0];
                var currentRightX = vb[2];
                var currentCenterX = (vb[0] + vb[2]) / 2;

                // Horizontal delta
                var deltaX = 0;
                if (applyHAlign) {
                    if (hAlign === "center") {
                        deltaX = cellCenterX - currentCenterX;
                    } else if (hAlign === "right") {
                        deltaX = cellRight - currentRightX;
                    } else {
                        deltaX = cellLeft - currentLeftX;
                    }
                    item.left = item.left + deltaX;
                }

                // Cell Y range (top -> bottom)
                var cellTop = currentY;
                var cellBottom = currentY - itemHeight;
                var cellCenterY = (cellTop + cellBottom) / 2;

                // Current anchors (Y)
                var currentTop = vb[1];
                var currentBottom = vb[3];
                var currentCenterY = (vb[1] + vb[3]) / 2;

                // Vertical delta (Top)
                var deltaY = cellTop - currentTop;
                item.top = item.top + deltaY;

                // Advance Y (downwards)
                if (r < itemsPerCol - 1) {
                    currentY -= itemHeight + vMarginPt;
                }
            }
        }

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
    function updatePreview() {
        // Roll back previous preview step(s)
        previewMgr.rollback();

        // Apply current preference for bounds calculation (this is not undoable, so we restore only on Cancel)
        app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value);

        // Apply as one preview step
        previewMgr.addStep(function () {
            applyLayoutToSelection();
        });
    }

    // N/L/C/R で揃えラジオを切替（dlg + 入力欄の両方に付与して取りこぼし防止）
    addAlignKeyHandler(dlg, rbHNone, rbHLeft, rbHCenter, rbHRight, updatePreview);
    addAlignKeyHandler(hMarginInput, rbHNone, rbHLeft, rbHCenter, rbHRight, updatePreview);

    updatePreview();
    randomCheckbox.onClick = function () {
        randomOrderCache = null;
        updatePreview();
    };

    /* ダイアログを開く前に 垂直方向（1段目）をアクティブにする / Activate vertical spacing (1st field) before showing dialog */
    hMarginInput.active = true;
    var dlgResult = dlg.show();
    saveDialogPosition(dlg.location);
    if (dlgResult !== 1) {
        // Cancel: roll back preview changes and restore preference
        previewMgr.rollback();
        app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
        app.redraw();
        return null;
    }

    // 垂直方向のみ / Vertical only
    var vMarginValue = parseFloat(hMarginInput.text);
    if (isNaN(vMarginValue)) vMarginValue = 0;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptFactor = getPtFactorFromUnitCode(unitCode);
    var hMarginPt = 0;
    var vMarginPt = vMarginValue * ptFactor;

    var hAlign = "left";
    if (rbHNone.value) hAlign = "none";
    else if (rbHCenter.value) hAlign = "center";
    else if (rbHRight.value) hAlign = "right";

    // Confirm as a single undoable action: rollback preview then run once
    previewMgr.confirm(function () {
        // keep the current preference value on OK (existing behavior effectively leaves it as-is)
        app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value);
        applyLayoutToSelection();
        app.redraw();
    });

    // Return values are kept for compatibility (main() currently just exits)
    return {
        mode: "vertical",
        hMargin: hMarginPt,
        vMargin: vMarginPt,
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