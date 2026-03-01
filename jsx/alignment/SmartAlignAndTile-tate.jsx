#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AlignAndDistributeTate.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignAndTile.jsx

### 概要：

- 更新日：2026-02-26
- 選択したオブジェクトを縦方向に整列し、指定した間隔と横方向の数（列数）で再配置するスクリプト。
- プレビュー時に境界線を含むオプション、ランダム配置、単位の自動取得、上下キーでの数値変更に対応。
- プレビュー時にUndo履歴を汚さないように管理し、OK時は1回のUndoで取り消せるように確定します。

### 主な機能：

- 縦方向整列と再配置
- 列数（横方向の数）指定
- ランダム配置オプション
- プレビュー時の境界含む切替
- 単位自動対応
- キーボードで間隔・列数調整
- Undoを汚さないプレビューと一括取り消し（1回のUndo）

### 処理の流れ：

- オブジェクト選択確認
- ダイアログ表示（各オプション設定）
- プレビュー更新
- 実行時に配置確定

### オリジナルアイデア

John Wundes 
Distribute Stacked Objects v1.1
https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx

Gorolib Design
https://gorolib.blog.jp/archives/77282974.html

### 更新履歴：
- v1.8 (20260226) : 縦方向前提（列数指定）にロジックとUIを変更、変数/関数名を整理
- v1.7 (20260119) : プレビュー時にUndo履歴を汚さないように管理し、OK時は1回のUndoで取り消せるように確定
- v1.6 (20250809) : 「プレビュー境界を使用」をOFFのとき geometricBounds を使用するように調整
- v1.0 (20250716) : 初期バージョン
- v1.1 (20250717) : 安定性改善、行数ロジック修正
- v1.2 (20250718) : コメント整理、ローカライズ統一、ランダム基準位置補正改善
- v1.3 (20250801) : グリッド機能の追加、ガターを縦横個別に設定
- v1.4 (20250801) : ローカライズを調整
- v1.5 (20250802) : 横／縦の連動機能を追加、プレビュー境界を使用のロジックを調整
- v1.6 (20250809) : 「プレビュー境界を使用」をOFFのとき geometricBounds を使用するように調整

---

### Script Name:

AlignAndDistributeTate.jsx

### Readme (GitHub):

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartAlignAndTile.jsx

### Overview:

- Updated: 2026-02-26
- Arrange selected objects vertically and re-distribute by specified spacing and number of columns.
- Supports preview bounds option, random arrangement, auto unit detection, and keyboard adjustments.
- Uses an undo-safe preview workflow and confirms the result so it can be undone with a single Undo step.

### Main Features:

- Vertical arrangement and redistribution
- Column count (horizontal count) specification
- Random arrangement option
- Toggle including preview bounds
- Automatic unit detection
- Keyboard adjustment for spacing and columns
- Undo-safe preview and single-step Undo on confirm

### Workflow:

- Check object selection
- Show dialog and set options
- Preview update
- Confirm to apply

### Update History:
- v1.8 (2026-02-26): Switched logic/UI to vertical-first layout (columns), refactored names
- v1.7 (2026-01-19): Undo-safe preview management and single-step Undo on confirm.
- v1.6 (2025-08-09): When "Use preview bounds" is OFF, use geometricBounds (OFF = geometric, ON = visible).
- v1.0 (2025-07-16): Initial version
- v1.1 (2025-07-17): Stability improvements, row logic fix
- v1.2 (2025-07-18): Comment refinement, localization update, improved random positioning correction
- v1.3 (2025-08-01): Added grid feature, separate gutter settings for horizontal and vertical spacing
- v1.4 (2025-08-01): Adjusted localization
- v1.5 (2025-08-02): Added horizontal/vertical linking feature, adjusted preview bounds logic
- v1.6 (2025-08-09): Adjusted to use geometricBounds when "Use preview bounds" is OFF
*/

/* バージョン変数を追加 / Script version variable */
var SCRIPT_VERSION = "v1.8";

/* ダイアログ位置と外観変数 / Dialog position and appearance variables */
var offsetX = 300;
var offsetY = 0;
var dialogOpacity = 0.97;

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* LABELS 定義（ダイアログUIの順序に合わせて並べ替え）/ Label definitions (ordered to match dialog UI) */
var LABELS = {
    dialogTitle: {
        ja: "整列と分布（縦）" + SCRIPT_VERSION,
        en: "Align & Distribute (Vertical) " + SCRIPT_VERSION
    },
    cols: {
        ja: "列数:",
        en: "Cols:"
    },
    hMargin: {
        ja: "垂直方向:",
        en: "V:"
    },
    vMargin: {
        ja: "水平方向:",
        en: "H:"
    },
    useBounds: {
        ja: "プレビュー境界を使用",
        en: "Use preview bounds"
    },
    grid: {
        ja: "グリッド",
        en: "Grid"
    },
    random: {
        ja: "ランダム",
        en: "Random"
    },
    alignPanel: {
        ja: "揃え",
        en: "Align"
    },
    vAlignTop: {
        ja: "上",
        en: "Top"
    },
    vAlignMiddle: {
        ja: "中央",
        en: "Middle"
    },
    vAlignBottom: {
        ja: "下",
        en: "Bottom"
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
    this.addStep = function(func) {
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
    this.rollback = function() {
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
    this.confirm = function(finalAction) {
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

/* ダイアログ位置をオフセットするヘルパー / Helper to shift dialog position by offsetX/offsetY */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/* ダイアログの不透明度を設定するヘルパー / Helper to set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
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
    shiftDialogPosition(dlg, offsetX, offsetY);


    // Undo-safe preview manager
    var previewMgr = new PreviewManager();

    // Preserve current preference so cancel can restore it
    var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");


    /* 間隔パネル / Spacing panel */
    var unit = getCurrentUnitLabel();
    var spacingPanel = dlg.add("panel", undefined, (lang === "ja" ? "間隔 (" + unit + ")" : "Spacing (" + unit + ")"));
    spacingPanel.orientation = "column";
    spacingPanel.alignChildren = ["fill", "top"];
    spacingPanel.margins = [15, 20, 15, 10];

    // 上段：入力（左）＋連動（右） / Top row: inputs (left) + link (right)
    var spacingTopRow = spacingPanel.add("group");
    spacingTopRow.orientation = "row";
    spacingTopRow.alignChildren = ["fill", "top"];
    spacingTopRow.alignment = ["fill", "top"];

    // 左カラム (入力グループ) / Left column (input groups)
    var spacingLeftGroup = spacingTopRow.add("group");
    spacingLeftGroup.orientation = "column";
    spacingLeftGroup.alignChildren = ["left", "center"];

    // 垂直方向マージン入力（UI 1段目）
    var hMarginGroup = spacingLeftGroup.add("group");
    hMarginGroup.orientation = "row";
    hMarginGroup.alignChildren = ["left", "center"];
    var hMarginLabel = hMarginGroup.add("statictext", undefined, LABELS.hMargin[lang]);
    var hMarginInput = hMarginGroup.add("edittext", undefined, "0");
    hMarginInput.characters = 3;
    changeValueByArrowKey(hMarginInput, true, function() {
        if (linkCheckbox.value) {
            vMarginInput.text = hMarginInput.text;
        }
        updatePreview();
    });
    // 単位ラベルは削除

    // 水平方向マージン入力（UI 2段目）
    var vMarginGroup = spacingLeftGroup.add("group");
    vMarginGroup.orientation = "row";
    vMarginGroup.alignChildren = ["left", "center"];
    var vMarginLabel = vMarginGroup.add("statictext", undefined, LABELS.vMargin[lang]);
    var vMarginInput = vMarginGroup.add("edittext", undefined, "0");
    vMarginInput.characters = 3;
    changeValueByArrowKey(vMarginInput, true, updatePreview);
    // 単位ラベルは削除

    // 右カラム (チェックボックス) / Right column (checkbox)
    var spacingRightWrapper = spacingTopRow.add("group");
    spacingRightWrapper.alignment = ["fill", "fill"];
    spacingRightWrapper.alignChildren = ["left", "center"];

    var spacingRightGroup = spacingRightWrapper.add("group");
    spacingRightGroup.orientation = "column";
    spacingRightGroup.alignChildren = ["left", "center"];

    var linkCheckbox = spacingRightGroup.add("checkbox", undefined, "連動");
    linkCheckbox.value = true; // ダイアログを開いたときにON

    // 初期状態で縦入力をディムし、横の値を同期
    vMarginInput.enabled = false;
    vMarginInput.text = hMarginInput.text;

    // 横・縦間隔の連動とプレビュー更新
    function syncMarginsAndPreview() {
        if (linkCheckbox.value) {
            // 連動ONのとき、2つの値を同期 / Sync the two values when linked
            vMarginInput.text = hMarginInput.text;
        }
        updatePreview();
    }

    // 「連動」チェックボックスの動作 / Link checkbox behavior
    linkCheckbox.onClick = function() {
        vMarginInput.enabled = !linkCheckbox.value;
        if (linkCheckbox.value) {
            vMarginInput.text = hMarginInput.text;
        }
        updatePreview();
    };

    // 連動ONのとき、2つの値を同期 / Sync the two values when linked
    hMarginInput.onChanging = syncMarginsAndPreview;
    hMarginInput.onChange = syncMarginsAndPreview;

    /* 揃えパネル / Align panel */
    var alignPanel = dlg.add("panel", undefined, LABELS.alignPanel[lang]);
    alignPanel.orientation = "column";
    alignPanel.alignChildren = ["left", "center"];
    alignPanel.margins = [15, 20, 15, 10];

    var applyAlignCheckbox = alignPanel.add("checkbox", undefined, (lang === "ja" ? "揃えを適用" : "Apply align"));
    applyAlignCheckbox.value = false; // デフォルトOFF

    /* 左右方向揃えグループ（基本はこちらを使用） / Horizontal align group (primary) */
    var hAlignGroup = alignPanel.add("group");
    hAlignGroup.orientation = "row";
    hAlignGroup.alignChildren = ["left", "center"];
    var rbHLeft = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignLeft[lang]);
    var rbHCenter = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignCenter[lang]);
    var rbHRight = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignRight[lang]);
    rbHLeft.value = true; /* デフォルトを「左」に / Default to Left */

    function syncHAlignEnable() {
        rbHLeft.enabled = rbHCenter.enabled = rbHRight.enabled = applyAlignCheckbox.value;
    }
    applyAlignCheckbox.onClick = function() {
        syncHAlignEnable();
        updatePreview();
    };
    syncHAlignEnable();

    rbHLeft.onClick = function() { updatePreview(); };
    rbHCenter.onClick = function() { updatePreview(); };
    rbHRight.onClick = function() { updatePreview(); };

    /* グリッドパネル（揃えの下に配置） / Grid panel (placed under Align) */
    var gridPanel = dlg.add("panel", undefined, LABELS.grid[lang]);
    gridPanel.orientation = "column";
    gridPanel.alignChildren = ["left", "center"];
    gridPanel.margins = [15, 20, 15, 10];

    // 上段：グリッドチェック + 列数（右） / Top row: Grid checkbox + Cols (right)
    var gridTopRow = gridPanel.add("group");
    gridTopRow.orientation = "row";
    gridTopRow.alignChildren = ["left", "center"];
    gridTopRow.alignment = ["fill", "center"];

    var gridCheckbox = gridTopRow.add("checkbox", undefined, "");
    gridCheckbox.value = false;

    var colsGroup = gridTopRow.add("group");
    colsGroup.orientation = "row";
    colsGroup.alignChildren = ["left", "center"];

    var colsLabel = colsGroup.add("statictext", undefined, LABELS.cols[lang]);
    var colsInput = colsGroup.add("edittext", undefined, "1");
    colsInput.characters = 3;
    // Helper to enable/disable cols UI by grid state
    function syncColsEnableByGrid(isGrid) {
        colsLabel.enabled = isGrid;
        colsInput.enabled = isGrid;
    }
    changeValueByArrowKey(colsInput, true, updatePreview);

    /* 上下方向揃え（グリッド時のみ使用） / Vertical align (grid-only) */
    var vAlignGroup = gridPanel.add("group");
    vAlignGroup.orientation = "row";
    vAlignGroup.alignChildren = ["left", "center"];
    var rbTop = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignTop[lang]);
    var rbMiddle = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignMiddle[lang]);
    var rbBottom = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignBottom[lang]);
    rbTop.value = true;

    rbTop.onClick = function() { updatePreview(); };
    rbMiddle.onClick = function() { updatePreview(); };
    rbBottom.onClick = function() { updatePreview(); };

    // グリッドOFF時は上下揃えを使わないため無効化 / Disable vertical align when grid is OFF
    function syncAlignEnableByGrid(isGrid) {
        rbTop.enabled = rbMiddle.enabled = rbBottom.enabled = isGrid;
    }

    gridCheckbox.onClick = function() {
        syncAlignEnableByGrid(gridCheckbox.value);
        syncColsEnableByGrid(gridCheckbox.value);
        if (gridCheckbox.value) {
            rbMiddle.value = true; // グリッド時は上下を中央に
            rbHCenter.value = true; // グリッド時は左右も中央に
            colsInput.text = "2"; // グリッドON時は列数を2に
            applyAlignCheckbox.value = true; // グリッドON時は「揃えを適用」をONに
            syncHAlignEnable(); // 左/中央/右を有効化
        } else {
            rbTop.value = true; // グリッドOFF時は上下揃えは使わない（値は上に戻しておく）
            colsInput.text = "1"; // グリッドOFF時は列数を1に
        }
        updatePreview();
    };

    // 初期状態を同期
    syncAlignEnableByGrid(gridCheckbox.value);
    syncColsEnableByGrid(gridCheckbox.value);

    // プレビュー境界チェック（間隔パネル最下部） / Preview bounds (bottom of spacing panel)
    var boundsRow = spacingPanel.add("group");
    boundsRow.orientation = "row";
    boundsRow.alignChildren = ["left", "center"];
    boundsRow.alignment = ["fill", "center"];

    var boundsCheckbox = boundsRow.add("checkbox", undefined, LABELS.useBounds[lang]);
    boundsCheckbox.value = true;
    boundsCheckbox.onClick = function() {
        updatePreview();
    };

    /* ランダム配置チェックボックス / Random arrangement checkbox */
    var randomCheckbox = dlg.add("checkbox", undefined, LABELS.random[lang]);
    randomCheckbox.value = false;

    /* ボタン配置グループ / Button group */
    var buttonGroup2 = dlg.add("group");
    buttonGroup2.alignment = "right";
    buttonGroup2.alignChildren = ["right", "center"];
    var cancelButton = buttonGroup2.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var okButton = buttonGroup2.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    // Keep a snapshot of the selection reference array (used for preview/final)
    var originalSelection = activeDocument.selection.slice();

    function applyLayoutToSelection() {
        // Guard
        if (!originalSelection || originalSelection.length === 0) return;

        // UI is swapped: first field is vertical, second is horizontal
        var vMarginValue = parseFloat(hMarginInput.text);
        if (isNaN(vMarginValue)) vMarginValue = 0;
        var hMarginValue = parseFloat(vMarginInput.text);
        if (isNaN(hMarginValue)) hMarginValue = 0;

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var ptFactor = getPtFactorFromUnitCode(unitCode);
        var hMarginPt = hMarginValue * ptFactor;
        var vMarginPt = vMarginValue * ptFactor;

        var colsValue = parseInt(colsInput.text, 10);
        if (isNaN(colsValue) || colsValue < 1) colsValue = 1;

        // Horizontal align (primary)
        var applyHAlign = applyAlignCheckbox.value;
        var hAlign = "left";
        if (rbHCenter.value) hAlign = "center";
        else if (rbHRight.value) hAlign = "right";

        // Vertical align is used only in Grid mode
        var vAlign = "top";
        if (gridCheckbox.value) {
            if (rbMiddle.value) vAlign = "middle";
            else if (rbBottom.value) vAlign = "bottom";
        }

        // Sort items (or shuffle)
        var sortedItems;
        var baseLeft = null;
        var baseTop = null;

        if (randomCheckbox.value) {
            sortedItems = originalSelection.slice();
            for (var i = sortedItems.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var temp = sortedItems[i];
                sortedItems[i] = sortedItems[j];
                sortedItems[j] = temp;
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
            // Vertical-first: sort by Y (top descending), then by X (left ascending)
            sortedItems = sortByY(originalSelection);
        }

        var cols = colsValue;
        if (cols < 1) cols = 1;
        var itemsPerCol = Math.ceil(sortedItems.length / cols);

        // Use selected bounds type (visible/geometric)
        var startBounds = getItemBounds(sortedItems[0], boundsCheckbox.value);
        var startLeftBound = startBounds[0];
        var startTopBound = startBounds[1];
        var startRightBound = startBounds[2];
        var startBottomBound = startBounds[3];
        var startX = startLeftBound;
        var startY = startTopBound;

        // Reference cell size (max w/h) for grid mode
        var refHeight = 0;
        var refWidth = 0;
        for (var m = 0; m < sortedItems.length; m++) {
            var vb0 = getItemBounds(sortedItems[m], boundsCheckbox.value);
            var w0 = vb0[2] - vb0[0];
            var h0 = vb0[1] - vb0[3];
            if (h0 > refHeight) refHeight = h0;
            if (w0 > refWidth) refWidth = w0;
        }

        var index = 0;
        for (var c = 0; c < cols; c++) {
            var currentY = startY;
            // Column X is based on the first item in the column in non-grid mode.
            // We compute cell width per item; but for advancing columns, use refWidth in grid mode.
            var colX = startX;

            // Compute colX by looking at the first item that will be placed in this column (if exists)
            if (!gridCheckbox.value) {
                var probeIndex = c * itemsPerCol;
                if (probeIndex < sortedItems.length) {
                    var vbProbe = getItemBounds(sortedItems[probeIndex], boundsCheckbox.value);
                    colX = vbProbe[0];
                }
                // Still keep the first column aligned to startX
                if (c === 0) colX = startX;
            }

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
                var cellBottom = currentY - (gridCheckbox.value ? refHeight : itemHeight);
                var cellCenterY = (cellTop + cellBottom) / 2;

                // Current anchors (Y)
                var currentTop = vb[1];
                var currentBottom = vb[3];
                var currentCenterY = (vb[1] + vb[3]) / 2;

                // Vertical delta
                var deltaY = 0;
                if (vAlign === "middle") {
                    deltaY = cellCenterY - currentCenterY;
                } else if (vAlign === "bottom") {
                    deltaY = cellBottom - currentBottom;
                } else {
                    deltaY = cellTop - currentTop;
                }
                item.top = item.top + deltaY;

                // Advance Y (downwards)
                if (r < itemsPerCol - 1) {
                    currentY -= (gridCheckbox.value ? refHeight : itemHeight) + vMarginPt;
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
        previewMgr.addStep(function() {
            applyLayoutToSelection();
        });
    }

    updatePreview();
    // syncAlignEnableByGrid(gridCheckbox.value); // already called above

    // (hMarginInput.onChanging is handled above for syncing)
    vMarginInput.onChanging = function() {
        updatePreview();
    };
    randomCheckbox.onClick = function() {
        updatePreview();
    };

    /* ダイアログを開く前に 垂直方向（1段目）をアクティブにする / Activate vertical spacing (1st field) before showing dialog */
    hMarginInput.active = true;
    if (dlg.show() !== 1) {
        // Cancel: roll back preview changes and restore preference
        previewMgr.rollback();
        app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
        app.redraw();
        return null;
    }

    // UI is swapped: first field is vertical, second is horizontal
    var vMarginValue = parseFloat(hMarginInput.text);
    if (isNaN(vMarginValue)) vMarginValue = 0;
    var hMarginValue = parseFloat(vMarginInput.text);
    if (isNaN(hMarginValue)) hMarginValue = 0;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptFactor = getPtFactorFromUnitCode(unitCode);
    var hMarginPt = hMarginValue * ptFactor;
    var vMarginPt = vMarginValue * ptFactor;

    var colsValue = parseInt(colsInput.text, 10);
    if (isNaN(colsValue) || colsValue < 1) colsValue = 1;

    var align = "top";
    if (rbMiddle.value) align = "middle";
    else if (rbBottom.value) align = "bottom";

    var hAlign = "left";
    if (rbHCenter.value) hAlign = "center";
    else if (rbHRight.value) hAlign = "right";

    // Confirm as a single undoable action: rollback preview then run once
    previewMgr.confirm(function() {
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
        cols: colsValue,
        random: randomCheckbox.value,
        vAlign: align,
        grid: gridCheckbox.value,
        hAlign: hAlign
    };
}

/* 編集テキストで上下キーによる数値変更を有効化 / Enable up/down arrow key increment/decrement on edittext inputs */
function changeValueByArrowKey(editText, allowNegative, onUpdate) {
    editText.addEventListener("keydown", function(event) {
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

/* 位置を元に戻す / Reset positions to original */
function resetPositions(items, positions) {
    for (var i = 0; i < items.length; i++) {
        items[i].left = positions[i][0];
        items[i].top = positions[i][1];
    }
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

/* Y座標でソート（上から下、同じなら左から右） / Sort items by Y coordinate (top to bottom, then left) */
function sortByY(items) {
    var copiedItems = items.slice();
    copiedItems.sort(function(a, b) {
        if (a.top !== b.top) return b.top - a.top; // higher top first
        return a.left - b.left;
    });
    return copiedItems;
}