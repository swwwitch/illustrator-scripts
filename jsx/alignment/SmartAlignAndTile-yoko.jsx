#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

重なって配置されたオブジェクトを、横方向へ指定した間隔で並べ直します。

詳細は README を参照してください。

### Overview

Redistributes stacked objects along the horizontal axis at the spacing you specify.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignAndTile-yoko";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-01-19";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile-yoko.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile-yoko.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
        ja: "整列と分布（グリッド対応）" + SCRIPT_VERSION,
        en: "Align & Distribute (Horizontal) " + SCRIPT_VERSION
    },
    rows: {
        ja: "行数:",
        en: "Rows:"
    },
    hMargin: {
        ja: "横:",
        en: "H:"
    },
    vMargin: {
        ja: "縦:",
        en: "V:"
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

    /* 行数入力UI: ラベルとテキストフィールドを横並びで配置 / Rows input UI: label and field side by side */
    var rowsGroup = dlg.add("group");
    rowsGroup.orientation = "row";
    rowsGroup.alignChildren = ["center", "center"];
    rowsGroup.alignment = ["fill", "center"];

    var rowsLabel = rowsGroup.add("statictext", undefined, LABELS.rows[lang]);
    var rowsInput = rowsGroup.add("edittext", undefined, "1");
    rowsInput.characters = 3;
    changeValueByArrowKey(rowsInput, true, updatePreview);

    /* 間隔パネル / Spacing panel */
    var unit = getCurrentUnitLabel();
    var spacingPanel = dlg.add("panel", undefined, "間隔 (" + unit + ")");
    spacingPanel.orientation = "row";
    spacingPanel.alignChildren = ["fill", "top"];
    spacingPanel.margins = [15, 20, 15, 10];

    // 左カラム (入力グループ) / Left column (input groups)
    var spacingLeftGroup = spacingPanel.add("group");
    spacingLeftGroup.orientation = "column";
    spacingLeftGroup.alignChildren = ["left", "center"];

    // 横方向マージン入力
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

    // 縦方向マージン入力
    var vMarginGroup = spacingLeftGroup.add("group");
    vMarginGroup.orientation = "row";
    vMarginGroup.alignChildren = ["left", "center"];
    var vMarginLabel = vMarginGroup.add("statictext", undefined, LABELS.vMargin[lang]);
    var vMarginInput = vMarginGroup.add("edittext", undefined, "0");
    vMarginInput.characters = 3;
    changeValueByArrowKey(vMarginInput, true, updatePreview);
    // 単位ラベルは削除

    // 右カラム (チェックボックス) / Right column (checkbox)
    var spacingRightWrapper = spacingPanel.add("group");
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

    // 横方向入力変更時に縦方向へ反映（連動がONのとき） / Sync V to H when linked
    hMarginInput.onChanging = syncMarginsAndPreview;
    hMarginInput.onChange = syncMarginsAndPreview;

    /* 揃えパネル / Align panel */
    var alignPanel = dlg.add("panel", undefined, LABELS.alignPanel[lang]);
    alignPanel.orientation = "column";
    alignPanel.alignChildren = ["left", "center"];
    alignPanel.margins = [15, 20, 15, 10];

    /* 上下方向揃えグループ / Vertical align group */
    var vAlignGroup = alignPanel.add("group");
    vAlignGroup.orientation = "row";
    vAlignGroup.alignChildren = ["left", "center"];
    var rbTop = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignTop[lang]);
    var rbMiddle = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignMiddle[lang]);
    var rbBottom = vAlignGroup.add("radiobutton", undefined, LABELS.vAlignBottom[lang]);
    rbTop.value = true; /* デフォルトを「上」に / Default to Top */

    rbTop.onClick = function() {
        updatePreview();
    };
    rbMiddle.onClick = function() {
        updatePreview();
    };
    rbBottom.onClick = function() {
        updatePreview();
    };

    /* 左右方向揃えグループ / Horizontal align group */
    var hAlignGroup = alignPanel.add("group");
    hAlignGroup.orientation = "row";
    hAlignGroup.alignChildren = ["left", "center"];
    var rbHLeft = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignLeft[lang]);
    var rbHCenter = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignCenter[lang]);
    var rbHRight = hAlignGroup.add("radiobutton", undefined, LABELS.hAlignRight[lang]);
    rbHLeft.value = true; /* デフォルトを「左」に / Default to Left */
    rbHLeft.onClick = function() {
        updatePreview();
    };
    rbHCenter.onClick = function() {
        updatePreview();
    };
    rbHRight.onClick = function() {
        updatePreview();
    };

    /* プレビュー境界チェックボックス / Preview bounds checkbox */
    var boundsCheckbox = dlg.add("checkbox", undefined, LABELS.useBounds[lang]);
    boundsCheckbox.value = true;
    boundsCheckbox.onClick = function() {
        updatePreview();
    };

    /* グリッドチェックボックス / Grid checkbox */
    var gridCheckbox = dlg.add("checkbox", undefined, LABELS.grid[lang]);
    gridCheckbox.value = false;
    gridCheckbox.onClick = function() {
        rbHLeft.enabled = rbHCenter.enabled = rbHRight.enabled = gridCheckbox.value;
        if (gridCheckbox.value) {
            rbMiddle.value = true; // 縦方向を中央に
            rbHCenter.value = true; // 横方向を中央に
        }
        updatePreview();
    };
    // 初期状態を同期
    rbHLeft.enabled = rbHCenter.enabled = rbHRight.enabled = gridCheckbox.value;

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

    /* 実際の配置処理（プレビュー／確定共通）/ Layout function used by both preview and final */
    function applyLayoutToSelection() {
        // Guard
        if (!originalSelection || originalSelection.length === 0) return;

        var hMarginValue = parseFloat(hMarginInput.text);
        if (isNaN(hMarginValue)) hMarginValue = 0;
        var vMarginValue = parseFloat(vMarginInput.text);
        if (isNaN(vMarginValue)) vMarginValue = 0;

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var ptFactor = getPtFactorFromUnitCode(unitCode);
        var hMarginPt = hMarginValue * ptFactor;
        var vMarginPt = vMarginValue * ptFactor;

        var rowsValue = parseInt(rowsInput.text, 10);
        if (isNaN(rowsValue) || rowsValue < 1) rowsValue = 1;

        var align = "top";
        if (rbMiddle.value) align = "middle";
        else if (rbBottom.value) align = "bottom";

        var hAlign = "left";
        if (rbHCenter.value) hAlign = "center";
        else if (rbHRight.value) hAlign = "right";

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
            sortedItems = sortByX(originalSelection);
        }

        var rows = rowsValue;
        if (rows < 1) rows = 1;
        var itemsPerRow = Math.ceil(sortedItems.length / rows);

        // Use selected bounds type (visible/geometric)
        var startBounds = getItemBounds(sortedItems[0], boundsCheckbox.value);
        var startLeftBound = startBounds[0];
        var startTopBound = startBounds[1];
        var startRightBound = startBounds[2];
        var startBottomBound = startBounds[3];
        var startX = startLeftBound;
        var startY = startTopBound;

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
        for (var r = 0; r < rows; r++) {
            var currentX = startX;
            for (var c = 0; c < itemsPerRow && index < sortedItems.length; c++, index++) {
                var item = sortedItems[index];
                if (!item) continue;

                var vb = getItemBounds(item, boundsCheckbox.value);
                var itemWidth = vb[2] - vb[0];
                var itemHeight = vb[1] - vb[3];

                // Cell X range
                var cellLeft = currentX;
                var cellRight = currentX + (gridCheckbox.value ? refWidth : itemWidth);
                var cellCenterX = (cellLeft + cellRight) / 2;

                // Current anchors
                var currentLeftX = vb[0];
                var currentRightX = vb[2];
                var currentCenterX = (vb[0] + vb[2]) / 2;

                // Horizontal delta
                var deltaX = 0;
                if (hAlign === "center") {
                    deltaX = cellCenterX - currentCenterX;
                } else if (hAlign === "right") {
                    deltaX = cellRight - currentRightX;
                } else {
                    deltaX = cellLeft - currentLeftX;
                }
                item.left = item.left + deltaX;

                // Cell Y positions
                var cellTop = startTopBound - (r * (refHeight + vMarginPt));
                var cellBottom = startBottomBound - (r * (refHeight + vMarginPt));
                var cellCenter = (cellTop + cellBottom) / 2;

                // Current Y anchors
                var currentTop = vb[1];
                var currentBottom = vb[3];
                var currentCenter = (vb[1] + vb[3]) / 2;

                // Vertical delta
                var deltaY = 0;
                if (align === "middle") {
                    deltaY = cellCenter - currentCenter;
                } else if (align === "bottom") {
                    deltaY = cellBottom - currentBottom;
                } else {
                    deltaY = cellTop - currentTop;
                }
                item.top = item.top + deltaY;

                // Advance X
                if (c < itemsPerRow - 1) {
                    if (gridCheckbox.value) {
                        currentX += refWidth + hMarginPt;
                    } else {
                        currentX += itemWidth + hMarginPt;
                    }
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

    // (hMarginInput.onChanging is handled above for syncing)
    vMarginInput.onChanging = function() {
        updatePreview();
    };
    randomCheckbox.onClick = function() {
        updatePreview();
    };

    /* ダイアログを開く前に rowsInput をアクティブにする / Activate rows input before showing dialog */
    rowsInput.active = true;
    if (dlg.show() !== 1) {
        // Cancel: roll back preview changes and restore preference
        previewMgr.rollback();
        app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
        app.redraw();
        return null;
    }

    var hMarginValue = parseFloat(hMarginInput.text);
    if (isNaN(hMarginValue)) hMarginValue = 0;
    var vMarginValue = parseFloat(vMarginInput.text);
    if (isNaN(vMarginValue)) vMarginValue = 0;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptFactor = getPtFactorFromUnitCode(unitCode);
    var hMarginPt = hMarginValue * ptFactor;
    var vMarginPt = vMarginValue * ptFactor;

    var rowsValue = parseInt(rowsInput.text, 10);
    if (isNaN(rowsValue) || rowsValue < 1) rowsValue = 1;

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
        mode: "horizontal",
        hMargin: hMarginPt,
        vMargin: vMarginPt,
        rows: rowsValue,
        random: randomCheckbox.value,
        align: align,
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

/* X座標でソート / Sort items by X coordinate */
function sortByX(items) {
    var copiedItems = items.slice();
    copiedItems.sort(function(a, b) {
        return a.left - b.left;
    });
    return copiedItems;
}
