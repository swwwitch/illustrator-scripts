#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択した2つのオブジェクトの位置を入れ替えます。
- 「中心位置を交換」：お互いの中心位置を交換します。
- 「両端の位置を保って交換」：左端・右端を固定したまま入れ替えます。
- 「見た目のサイズを基準にする」で、線幅や効果を含めた bounds を基準にできます。
- プレビューを見ながら基準を切り替えられ、キャンセルで元の位置に戻ります。

### 処理の流れ

1. 選択が2つの入れ替え可能なオブジェクトかを検証
2. ダイアログで基準（位置・サイズ）を選択
3. 元の位置を控えたうえで移動量を計算し、適用
4. OK で確定、キャンセルで復元

### Overview

- Swaps the positions of two selected objects.
- "Swap Center Positions": exchanges their center points.
- "Swap Keeping Outer Edges": swaps them while keeping the outer left and right edges fixed.
- "Use Visual Bounds" switches the reference to bounds that include strokes and effects.
- The reference can be changed with live preview, and Cancel restores the original positions.

### Process

1. Validate that exactly two swappable objects are selected
2. Choose the position and size references in the dialog
3. Snapshot the original positions, calculate the offsets, and apply them
4. Commit with OK, or restore with Cancel

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "SwapObjects";                  /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-04-06";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-07-21";                   /* 更新日 / last updated */

    // README (Japanese)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapObjects.md
    // README (English)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapObjects.md
    var SCRIPT_URL      = "https://note.com/dtp_tranist/n/na534a676fae2";                             /* 紹介記事 / article URL */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_USE_VISUAL_BOUNDS      = false;  /* 「見た目のサイズを基準にする」の初期値 / default for visual bounds */
    var DEFAULT_LIVE_PREVIEW_ENABLED   = false;  /* プレビューの初期値 / default for live preview */
    var EQUAL_WIDTH_TOLERANCE          = 0.01;   /* 幅を同値とみなす許容差 / tolerance for equal widths */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの位置を入れ替え", en: "Swap Object Positions" }
        },
        panel: {
            positionReference: { ja: "位置の基準", en: "Position Reference" },
            sizeReference: { ja: "サイズの基準", en: "Size Reference" }
        },
        radio: {
            swapCenters: { ja: "中心位置を交換", en: "Swap Center Positions" },
            keepOuterEdges: { ja: "両端の位置を保って交換", en: "Swap Keeping Outer Edges" }
        },
        checkbox: {
            useVisualBounds: {
                ja: "見た目のサイズを基準にする（線幅・効果を含む）",
                en: "Use Visual Bounds (Include Stroke and Effects)"
            },
            livePreview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            swapCenters: {
                ja: "お互いの中心が入れ替わります。\n幅が異なる場合、2つ並びの左端・右端の位置は変わります。",
                en: "The two objects exchange their center points.\nIf their widths differ, the outer left and right edges will shift."
            },
            keepOuterEdges: {
                ja: "左端と右端の位置を保ったまま入れ替えます。\n幅が異なっていても、2つ並び全体の占有幅は変わりません。",
                en: "Swaps the objects while keeping the outer left and right edges in place.\nThe total span stays the same even when their widths differ."
            },
            useVisualBounds: {
                ja: "オフ：パス本体のサイズ（geometric bounds）を基準にします。\nオン：線幅・効果を含む見た目のサイズ（visible bounds）を基準にします。",
                en: "Off: uses the path geometry (geometric bounds).\nOn: uses the appearance including strokes and effects (visible bounds)."
            },
            livePreview: {
                ja: "結果を画面で確認します。\nキャンセルすると元の位置に戻ります。",
                en: "Shows the result on the canvas.\nCancel restores the original positions."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントを開いてから実行してください",
                en: "Please open a document first."
            },
            selectTwoItems: {
                ja: "2つのオブジェクトを選択してください",
                en: "Please select two objects."
            },
            unsupportedItem: {
                ja: "対応していないオブジェクトが含まれています",
                en: "The selection contains unsupported objects."
            },
            lockedOrHidden: {
                ja: "ロックまたは非表示のオブジェクトは対象にできません",
                en: "Locked or hidden objects are not supported."
            }
        },
        error: {
            invalidBoundsMode: { ja: "不正な bounds mode: ", en: "Invalid bounds mode: " }
        }
    };

    /* 現在の言語のラベルを取得 / Get the label for the current language */
    function getLabel(category, key) {
        var labelEntry = LABELS[category][key];
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // 定数 / Constants
    // =========================================
    var BOUNDS_MODES = {
        GEOMETRIC: 'geometric',
        VISUAL: 'visual'
    };

    var REFERENCE_MODES = {
        SWAP_CENTERS: 'swapCenters',
        KEEP_OUTER_EDGES: 'keepOuterEdges'
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /* 前提条件を確認してダイアログを開く / Check prerequisites and open the dialog */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }

        var targetItems = app.activeDocument.selection;
        var validationResult = validateSwapSelection(targetItems);

        if (!validationResult.isValid) {
            alert(validationResult.message);
            return;
        }

        showSwapOptionsDialog(targetItems);
    }

    // =========================================
    // バリデーション / Validation
    // =========================================

    /* 選択が「2つの入れ替え可能なオブジェクト」か判定 / Verify the selection is two swappable objects */
    function validateSwapSelection(targetItems) {
        if (!targetItems || targetItems.length !== 2) {
            return { isValid: false, message: getLabel('alert', 'selectTwoItems') };
        }

        for (var i = 0; i < targetItems.length; i++) {
            if (!isSwappableItem(targetItems[i])) {
                return { isValid: false, message: getLabel('alert', 'unsupportedItem') };
            }
            if (isLockedOrHidden(targetItems[i])) {
                return { isValid: false, message: getLabel('alert', 'lockedOrHidden') };
            }
        }

        return { isValid: true, message: '' };
    }

    /* 移動でき、両方の bounds を取得できるオブジェクトか / Check the item can move and expose both bounds */
    function isSwappableItem(targetItem) {
        if (!targetItem || typeof targetItem.translate !== 'function') {
            return false;
        }

        try {
            return isBoundsArray(targetItem.geometricBounds) && isBoundsArray(targetItem.visibleBounds);
        } catch (e) {
            return false;
        }
    }

    /* bounds が4要素の配列か / Check the bounds array has four values */
    function isBoundsArray(bounds) {
        return !!bounds && bounds.length === 4;
    }

    /* 自身と祖先にロック・非表示がないか / Check the item and its ancestors for lock or hidden state */
    function isLockedOrHidden(targetItem) {
        for (var ancestor = targetItem; ancestor; ancestor = ancestor.parent) {
            /* 上位階層はプロパティ自体を持たず例外になり得るため一括で握る
               Ancestors may not expose these properties, so guard the whole check */
            try {
                if (ancestor.locked === true || ancestor.hidden === true || ancestor.visible === false) {
                    return true;
                }
            } catch (e) {}
        }

        return false;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ダイアログを組み立て、プレビューと確定処理を接続 / Build the dialog and wire preview and commit */
    function showSwapOptionsDialog(targetItems) {
        var dialogControls = buildSwapDialog();
        var swapController = createSwapController(targetItems);

        var initialReferenceMode = getInitialReferenceMode(
            swapController.getOriginalBoundsPair(BOUNDS_MODES.GEOMETRIC)
        );
        dialogControls.rdoSwapCenters.value = (initialReferenceMode === REFERENCE_MODES.SWAP_CENTERS);
        dialogControls.rdoKeepOuterEdges.value = (initialReferenceMode === REFERENCE_MODES.KEEP_OUTER_EDGES);

        /* ダイアログの現在値を設定オブジェクトにまとめる / Collect the current dialog state */
        function getSwapSettings() {
            return {
                boundsMode: dialogControls.chkUseVisualBounds.value ? BOUNDS_MODES.VISUAL : BOUNDS_MODES.GEOMETRIC,
                referenceMode: dialogControls.rdoKeepOuterEdges.value ? REFERENCE_MODES.KEEP_OUTER_EDGES : REFERENCE_MODES.SWAP_CENTERS
            };
        }

        /* プレビューの適用・解除を切り替えて再描画 / Toggle the preview and redraw */
        function refreshLivePreview() {
            var didChange = dialogControls.chkLivePreview.value ?
                swapController.applySwap(getSwapSettings()) :
                swapController.restoreOriginalPositions();

            if (didChange) {
                app.redraw();
            }
        }

        dialogControls.chkLivePreview.onClick = refreshLivePreview;
        dialogControls.rdoSwapCenters.onClick = refreshLivePreview;
        dialogControls.rdoKeepOuterEdges.onClick = refreshLivePreview;
        dialogControls.chkUseVisualBounds.onClick = refreshLivePreview;

        dialogControls.btnOk.onClick = function () {
            /* プレビュー適用済みならそのまま確定 / Keep the applied preview as the result */
            if (!swapController.isSwapApplied()) {
                swapController.applySwap(getSwapSettings());
            }
            app.redraw();
            dialogControls.dialogWindow.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.dialogWindow.close(0);
        };

        if (dialogControls.dialogWindow.show() !== 1) {
            if (swapController.restoreOriginalPositions()) {
                app.redraw();
            }
        }
    }

    /* ダイアログの各コントロールを生成して返す / Create the dialog controls and return them */
    function buildSwapDialog() {
        var dialogWindow = new Window('dialog', getLabel('dialog', 'title') + ' ' + SCRIPT_VERSION);
        setupWindow(dialogWindow);

        /* 位置の基準 / Position reference */
        var positionReferencePanel = dialogWindow.add('panel', undefined, getLabel('panel', 'positionReference'));
        setupPanel(positionReferencePanel, 6);

        var rdoSwapCenters = positionReferencePanel.add('radiobutton', undefined, getLabel('radio', 'swapCenters'));
        rdoSwapCenters.helpTip = getLabel('tooltip', 'swapCenters');

        var rdoKeepOuterEdges = positionReferencePanel.add('radiobutton', undefined, getLabel('radio', 'keepOuterEdges'));
        rdoKeepOuterEdges.helpTip = getLabel('tooltip', 'keepOuterEdges');

        /* サイズの基準 / Size reference */
        var sizeReferencePanel = dialogWindow.add('panel', undefined, getLabel('panel', 'sizeReference'));
        setupPanel(sizeReferencePanel, 6);

        var chkUseVisualBounds = sizeReferencePanel.add('checkbox', undefined, getLabel('checkbox', 'useVisualBounds'));
        chkUseVisualBounds.helpTip = getLabel('tooltip', 'useVisualBounds');
        chkUseVisualBounds.value = DEFAULT_USE_VISUAL_BOUNDS;

        /* ボタンエリアを左右分割で組み立てる。
           左：プレビュー、中央：伸縮スペーサー、右：キャンセル / OK。
           Build the footer split left and right: Preview on the left, a stretchable spacer
           in the middle, and Cancel / OK on the right. */
        var footerRowGroup = dialogWindow.add('group');
        footerRowGroup.orientation = 'row';
        footerRowGroup.margins = [10, 10, 10, 0];
        footerRowGroup.alignment = ['fill', 'bottom'];

        /* 左側グループ / Left-side group
           プレビューは押しっぱなしの切り替えなので、ボタンではなくチェックボックスにしている
           Preview is a sticky toggle, so it stays a checkbox rather than a button */
        var footerLeftGroup = footerRowGroup.add('group');
        footerLeftGroup.alignChildren = ['left', 'center'];
        var chkLivePreview = footerLeftGroup.add('checkbox', undefined, getLabel('checkbox', 'livePreview'));
        chkLivePreview.helpTip = getLabel('tooltip', 'livePreview');
        chkLivePreview.value = DEFAULT_LIVE_PREVIEW_ENABLED;

        /* スペーサー（伸縮） / Spacer (stretchable) */
        var footerSpacer = footerRowGroup.add('group');
        footerSpacer.alignment = ['fill', 'fill'];
        footerSpacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var footerRightGroup = footerRowGroup.add('group');
        footerRightGroup.alignChildren = ['right', 'center'];
        var btnCancel = footerRightGroup.add('button', undefined, getLabel('button', 'cancel'), { name: 'cancel' });
        var btnOk = footerRightGroup.add('button', undefined, getLabel('button', 'ok'), { name: 'ok' });

        return {
            dialogWindow: dialogWindow,
            rdoSwapCenters: rdoSwapCenters,
            rdoKeepOuterEdges: rdoKeepOuterEdges,
            chkUseVisualBounds: chkUseVisualBounds,
            chkLivePreview: chkLivePreview,
            btnCancel: btnCancel,
            btnOk: btnOk
        };
    }

    /* 幅が同じなら中心交換、違えば両端維持を初期選択 / Default to center swap for equal widths, keep-edges otherwise */
    function getInitialReferenceMode(geometricBoundsPair) {
        var widthA = getBoundsWidth(geometricBoundsPair[0]);
        var widthB = getBoundsWidth(geometricBoundsPair[1]);

        return isNearlyEqual(widthA, widthB) ? REFERENCE_MODES.SWAP_CENTERS : REFERENCE_MODES.KEEP_OUTER_EDGES;
    }

    /* 誤差を許容して同値とみなす / Compare two numbers with a small tolerance */
    function isNearlyEqual(valueA, valueB) {
        return Math.abs(valueA - valueB) < EQUAL_WIDTH_TOLERANCE;
    }

    // =========================================
    // 入れ替えの適用と復元 / Apply and restore the swap
    // =========================================

    /* 元の状態を保持し、適用・復元を行うオブジェクトを作る / Create a controller that holds the original state and applies or restores the swap */
    function createSwapController(targetItems) {
        /* 元の状態を先にスナップショット。適用は必ず元位置からやり直す
           Snapshot the original state; every apply starts over from it */
        var originalPositions = [getItemPosition(targetItems[0]), getItemPosition(targetItems[1])];
        var originalBoundsByMode = {};
        originalBoundsByMode[BOUNDS_MODES.GEOMETRIC] = snapshotBoundsPair(targetItems, BOUNDS_MODES.GEOMETRIC);
        originalBoundsByMode[BOUNDS_MODES.VISUAL] = snapshotBoundsPair(targetItems, BOUNDS_MODES.VISUAL);

        var swapApplied = false;

        /* 元の位置へ戻す。変化があれば true / Restore the original positions; true if anything moved */
        function restoreOriginalPositions() {
            if (!swapApplied) {
                return false;
            }

            setItemPosition(targetItems[0], originalPositions[0]);
            setItemPosition(targetItems[1], originalPositions[1]);
            swapApplied = false;
            return true;
        }

        return {
            isSwapApplied: function () {
                return swapApplied;
            },
            getOriginalBoundsPair: function (boundsMode) {
                return originalBoundsByMode[boundsMode];
            },
            restoreOriginalPositions: restoreOriginalPositions,
            /* 設定に従って入れ替えを適用 / Apply the swap for the given settings */
            applySwap: function (swapSettings) {
                restoreOriginalPositions();

                var boundsPair = originalBoundsByMode[swapSettings.boundsMode];
                var swapOffsets = calculateSwapOffsets(boundsPair[0], boundsPair[1], swapSettings.referenceMode);

                targetItems[0].translate(swapOffsets.itemA.x, swapOffsets.itemA.y);
                targetItems[1].translate(swapOffsets.itemB.x, swapOffsets.itemB.y);
                swapApplied = true;
                return true;
            }
        };
    }

    /* 位置を配列として控える / Copy the item position into a plain array */
    function getItemPosition(targetItem) {
        return [targetItem.position[0], targetItem.position[1]];
    }

    /* 控えた位置を書き戻す / Write a stored position back to the item */
    function setItemPosition(targetItem, storedPosition) {
        targetItem.position = [storedPosition[0], storedPosition[1]];
    }

    /* 2つのオブジェクトの bounds を複製して控える / Snapshot the bounds of both items */
    function snapshotBoundsPair(targetItems, boundsMode) {
        return [
            copyBounds(getBoundsByMode(targetItems[0], boundsMode)),
            copyBounds(getBoundsByMode(targetItems[1], boundsMode))
        ];
    }

    /* bounds 配列を複製 / Duplicate a bounds array */
    function copyBounds(bounds) {
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    /* 指定モードの bounds を取得 / Get the bounds for the given mode */
    function getBoundsByMode(targetItem, boundsMode) {
        if (boundsMode === BOUNDS_MODES.VISUAL) {
            return targetItem.visibleBounds;
        }
        if (boundsMode === BOUNDS_MODES.GEOMETRIC) {
            return targetItem.geometricBounds;
        }
        throw new Error(getLabel('error', 'invalidBoundsMode') + boundsMode);
    }

    // =========================================
    // 移動量の計算 / Calculate offsets
    // =========================================

    /* 基準モードに応じた移動量を返す / Return the offsets for the selected reference mode */
    function calculateSwapOffsets(boundsA, boundsB, referenceMode) {
        if (referenceMode === REFERENCE_MODES.KEEP_OUTER_EDGES) {
            return calculateKeepOuterEdgesOffsets(boundsA, boundsB);
        }
        return calculateSwapCentersOffsets(boundsA, boundsB);
    }

    /* 中心位置を交換 / Swap the center positions */
    function calculateSwapCentersOffsets(boundsA, boundsB) {
        var centerA = getBoundsCenter(boundsA);
        var centerB = getBoundsCenter(boundsB);
        var dx = centerB.x - centerA.x;
        var dy = centerB.y - centerA.y;

        return {
            itemA: { x: dx, y: dy },
            itemB: { x: -dx, y: -dy }
        };
    }

    /* 左端・右端を保ったまま入れ替え / Swap while keeping the outer left and right edges */
    function calculateKeepOuterEdgesOffsets(boundsA, boundsB) {
        var isItemAOnLeft = getBoundsCenter(boundsA).x <= getBoundsCenter(boundsB).x;
        var edgeSlots = createOuterEdgeSlots(
            isItemAOnLeft ? boundsA : boundsB,
            isItemAOnLeft ? boundsB : boundsA
        );

        return {
            itemA: getOffsetToEdgeSlot(boundsA, isItemAOnLeft ? edgeSlots.right : edgeSlots.left),
            itemB: getOffsetToEdgeSlot(boundsB, isItemAOnLeft ? edgeSlots.left : edgeSlots.right)
        };
    }

    /* 移動先となる左右のスロットを作る / Build the left and right destination slots */
    function createOuterEdgeSlots(leftBounds, rightBounds) {
        return {
            left: { edgeX: leftBounds[0], isRightEdge: false, centerY: getBoundsCenter(leftBounds).y },
            right: { edgeX: rightBounds[2], isRightEdge: true, centerY: getBoundsCenter(rightBounds).y }
        };
    }

    /* スロットに収めるための移動量を求める / Calculate the offset that seats the item into a slot */
    function getOffsetToEdgeSlot(bounds, targetSlot) {
        var targetLeft = targetSlot.isRightEdge ? (targetSlot.edgeX - getBoundsWidth(bounds)) : targetSlot.edgeX;
        var targetTop = targetSlot.centerY + (getBoundsHeight(bounds) / 2);

        return {
            x: targetLeft - bounds[0],
            y: targetTop - bounds[1]
        };
    }

    // =========================================
    // bounds のユーティリティ / Bounds utilities
    // =========================================

    /* bounds の中心座標 / Center point of the bounds */
    function getBoundsCenter(bounds) {
        return {
            x: (bounds[0] + bounds[2]) / 2,
            y: (bounds[1] + bounds[3]) / 2
        };
    }

    /* bounds の幅 / Width of the bounds */
    function getBoundsWidth(bounds) {
        return bounds[2] - bounds[0];
    }

    /* bounds の高さ（上が大きい座標系）/ Height of the bounds (y increases upward) */
    function getBoundsHeight(bounds) {
        return bounds[1] - bounds[3];
    }

    main();

})();
