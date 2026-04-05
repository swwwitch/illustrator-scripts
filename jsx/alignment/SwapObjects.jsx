#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


// =========================================
// スクリプト概要 / Script overview
// =========================================
// スクリプト名：オブジェクトの位置を入れ替え
// Script name: Swap Object Positions

// 概要：
// 選択した2つのオブジェクトの位置を入れ替えます。
// 「中央を基準に入れ替え」と「2つのオブジェクト全体の幅を保つ」の2つの基準に対応しています。
// 後者は2つのオブジェクト全体の幅を維持したまま配置を入れ替えます。
// サイズの基準は geometric / visible を切り替え可能です。
// ダイアログで設定変更ができ、プレビューにも対応しています。
// preview apply / preview remove / final apply を分離した設計になっています。
// プレビュー計算は初期状態の bounds を基準に行います。
// プレビュー中に基準を切り替えた場合は、現在の見た目上の左右関係を維持したまま再配置します。
// プレビュー表示中に［OK］を押した場合は、その見た目のまま確定します。
// ダイアログ初期表示時は、選択オブジェクトの幅を事前判定し、
// 同じ幅なら「中央を基準に入れ替え」、異なる幅なら「2つのオブジェクト全体の幅を保つ」を初期選択します。
// --- 簡易説明（ユーザー向け） ---
// 2つのオブジェクトの位置を直感的に入れ替えます。
// ・中央基準：お互いの中心位置を交換
// ・全体幅維持：2つ並びの外側サイズを崩さずに入れ替え
// プレビューを見ながら安全に調整できます。

// Summary:
// Swap the positions of two selected objects.
// Supports two modes: "Swap by Center" and "Preserve the Total Width of Both Objects".
// The latter keeps the total width occupied by both objects while swapping.
// Supports geometric or visible bounds.
// Includes a dialog UI with live preview.
// The design separates preview apply, preview remove, and final apply.
// Preview calculations are based on the original bounds.
// When switching modes during preview, the current visual left/right relationship is preserved.
// If OK is pressed while preview is visible, the current preview result is confirmed as-is.
// When the dialog opens, the script pre-checks the selected object widths.
// If the widths are the same, it defaults to "Swap by Center".
// If the widths differ, it defaults to "Preserve the Total Width of Both Objects".
// --- Simple Description (User-Friendly) ---
// Swap two objects intuitively.
// • Center mode: swap their center positions
// • Preserve width mode: keep the outer span while swapping
// Adjust safely with live preview.

// 更新日：2026-04-05

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================
var SCRIPT_VERSION = "v1.2.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "オブジェクトの位置を入れ替え",
        en: "Swap Object Positions"
    },
    labelReferenceMode: {
        ja: "位置の基準",
        en: "Reference Mode"
    },
    referenceCenter: {
        ja: "中央を基準に入れ替え",
        en: "Swap by Center"
    },
    referenceFixedEdges: {
        ja: "2つのオブジェクト全体の幅を保つ",
        en: "Preserve the Total Width of Both Objects"
    },
    referenceFixedEdgesSummary: {
        ja: "2つのオブジェクト全体の幅を固定",
        en: "Keep the Total Width of Both Objects Fixed"
    },
    labelBoundsMode: {
        ja: "サイズの基準",
        en: "Size Reference"
    },
    usePreviewBounds: {
        ja: "見た目のサイズを使用（線幅を含む）",
        en: "Use Visual Bounds (Include Stroke)"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertSelectTwoItems: {
        ja: "2つのオブジェクトを選択してください",
        en: "Please select two objects."
    },
    alertUnsupportedItem: {
        ja: "対応していないオブジェクトが含まれています",
        en: "The selection contains unsupported objects."
    },
    alertLockedOrHidden: {
        ja: "ロックまたは非表示のオブジェクトは対象にできません",
        en: "Locked or hidden objects are not supported."
    },
    errorInvalidBoundsMode: {
        ja: "不正な bounds mode: ",
        en: "Invalid bounds mode: "
    }
};

function L(key) {
    return LABELS[key][lang] || LABELS[key].en;
}

var BOUNDS_MODES = {
    GEOMETRIC: "geometric",
    VISIBLE: "visible"
};

var REFERENCE_MODES = {
    CENTER: 'center',
    FIXED_EDGES: 'fixedEdges'
};

var PREVIEW_RELAYOUT_POLICIES = {
    KEEP_CURRENT_VISUAL_ORDER: 'keepCurrentVisualOrder'
};

// =========================================
// メイン処理 / Main process
// =========================================
main(activeDocument);

function main(targetDoc) {
    var items = targetDoc.selection;
    var validationResult = validateSelection(items, BOUNDS_MODES.GEOMETRIC);

    if (!validationResult.isValid) {
        alert(validationResult.message);
        return;
    }

    showOptionsDialog(targetDoc, items);
}

// =========================================
// ダイアログを表示 / Show dialog
// =========================================
function showOptionsDialog(targetDoc, items) {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 10];

    var referencePanel = dlg.add('panel', undefined, L('labelReferenceMode'));
    referencePanel.orientation = 'column';
    referencePanel.alignChildren = ['left', 'top'];
    referencePanel.margins = [15, 20, 15, 10];

    var rdoReferenceCenter = referencePanel.add('radiobutton', undefined, L('referenceCenter'));
    var rdoReferenceFixedEdges = referencePanel.add('radiobutton', undefined, L('referenceFixedEdges'));

    var boundsPanel = dlg.add('panel', undefined, L('labelBoundsMode'));
    boundsPanel.orientation = 'column';
    boundsPanel.alignChildren = ['left', 'top'];
    boundsPanel.margins = [15, 20, 15, 10];

    var chkUsePreviewBounds = boundsPanel.add('checkbox', undefined, L('usePreviewBounds'));
    chkUsePreviewBounds.value = false;

    var bottomGroup = dlg.add('group');
    bottomGroup.orientation = 'row';
    bottomGroup.alignChildren = ['center', 'center'];

    var chkPreview = bottomGroup.add('checkbox', undefined, L('preview'));
    chkPreview.value = false;

    var spacer = bottomGroup.add('statictext', undefined, '');
    spacer.alignment = ['fill', 'fill'];

    var btnCancel = bottomGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var btnOk = bottomGroup.add('button', undefined, L('ok'), { name: 'ok' });

    var previewApplied = false;
    var isUpdatingPreview = false;
    var originalPositions = [
        getItemPosition(items[0]),
        getItemPosition(items[1])
    ];
    var originalBoundsByMode = {
        geometric: [
            copyBounds(items[0].geometricBounds),
            copyBounds(items[1].geometricBounds)
        ],
        visible: [
            copyBounds(items[0].visibleBounds),
            copyBounds(items[1].visibleBounds)
        ]
    };
    var initialReferenceMode = getInitialReferenceMode(originalBoundsByMode.geometric);
    rdoReferenceCenter.value = (initialReferenceMode === REFERENCE_MODES.CENTER);
    rdoReferenceFixedEdges.value = (initialReferenceMode === REFERENCE_MODES.FIXED_EDGES);

    function getSelectedReferenceMode() {
        if (rdoReferenceFixedEdges.value) {
            return REFERENCE_MODES.FIXED_EDGES;
        }
        return REFERENCE_MODES.CENTER;
    }

    function getSelectedBoundsMode() {
        return chkUsePreviewBounds.value ? BOUNDS_MODES.VISIBLE : BOUNDS_MODES.GEOMETRIC;
    }

    function getOriginalBoundsPair(boundsMode) {
        if (boundsMode === BOUNDS_MODES.VISIBLE) {
            return originalBoundsByMode.visible;
        }
        return originalBoundsByMode.geometric;
    }

    function getCurrentBoundsPair(boundsMode) {
        return [
            copyBounds(getBounds(items[0], boundsMode)),
            copyBounds(getBounds(items[1], boundsMode))
        ];
    }

    function getPreviewRelayoutPolicy() {
        return PREVIEW_RELAYOUT_POLICIES.KEEP_CURRENT_VISUAL_ORDER;
    }

    function getPreviewCalculationContext(settings) {
        return {
            baseBoundsPair: getOriginalBoundsPair(settings.boundsMode),
            visualBoundsPair: previewApplied ? getCurrentBoundsPair(settings.boundsMode) : getOriginalBoundsPair(settings.boundsMode),
            relayoutPolicy: getPreviewRelayoutPolicy()
        };
    }

    function getCurrentDialogSettings() {
        return {
            boundsMode: getSelectedBoundsMode(),
            referenceMode: getSelectedReferenceMode()
        };
    }

    function getSwapOffsetsForSettings(settings) {
        var calculationContext = getPreviewCalculationContext(settings);

        return calculateSwapOffsetsFromBounds(
            calculationContext.baseBoundsPair[0],
            calculationContext.baseBoundsPair[1],
            settings.referenceMode,
            calculationContext.visualBoundsPair[0],
            calculationContext.visualBoundsPair[1],
            calculationContext.relayoutPolicy
        );
    }

    function getInitialReferenceMode(originalGeometricBoundsPair) {
        var widthA = getWidth(originalGeometricBoundsPair[0]);
        var widthB = getWidth(originalGeometricBoundsPair[1]);

        if (areNearlyEqual(widthA, widthB)) {
            return REFERENCE_MODES.CENTER;
        }

        return REFERENCE_MODES.FIXED_EDGES;
    }

    function removePreview() {
        if (!previewApplied) {
            return false;
        }

        restoreItemPosition(items[0], originalPositions[0]);
        restoreItemPosition(items[1], originalPositions[1]);
        previewApplied = false;
        return true;
    }

    function applyPreview() {
        var settings = getCurrentDialogSettings();
        var swapOffsets = getSwapOffsetsForSettings(settings);

        removePreview();
        applySwapOffsets(items[0], items[1], swapOffsets);
        previewApplied = true;
        return true;
    }

    function applyFinal() {
        var settings = getCurrentDialogSettings();
        var swapOffsets = getSwapOffsetsForSettings(settings);

        removePreview();
        applySwapOffsets(items[0], items[1], swapOffsets);
        previewApplied = false;
        return true;
    }

    function updatePreview() {
        var didChange = false;

        if (isUpdatingPreview) {
            return;
        }

        isUpdatingPreview = true;

        try {
            if (chkPreview.value) {
                didChange = applyPreview();
            } else {
                didChange = removePreview();
            }

            if (didChange) {
                app.redraw();
            }
        } finally {
            isUpdatingPreview = false;
        }
    }

    chkPreview.onClick = updatePreview;
    rdoReferenceCenter.onClick = updatePreview;
    rdoReferenceFixedEdges.onClick = updatePreview;
    chkUsePreviewBounds.onClick = updatePreview;

    btnOk.onClick = function () {
        if (chkPreview.value && previewApplied) {
            previewApplied = false;
            app.redraw();
            dlg.close(1);
            return;
        }

        applyFinal();
        app.redraw();
        dlg.close(1);
    };

    btnCancel.onClick = function () {
        removePreview();
        dlg.close(0);
    };

    if (dlg.show() !== 1) {
        removePreview();
    }
}

// =========================================
// バリデーション / Validation
// =========================================
function validateSelection(items, boundsMode) {
    if (!items || items.length !== 2) {
        return {
            isValid: false,
            message: L('alertSelectTwoItems')
        };
    }

    for (var i = 0; i < items.length; i++) {
        if (!isSupportedItem(items[i], boundsMode)) {
            return {
                isValid: false,
                message: L('alertUnsupportedItem')
            };
        }

        if (isLockedOrHidden(items[i])) {
            return {
                isValid: false,
                message: L('alertLockedOrHidden')
            };
        }
    }

    return {
        isValid: true,
        message: ''
    };
}

function isSupportedItem(item, boundsMode) {
    if (!item) {
        return false;
    }

    if (typeof item.translate !== 'function') {
        return false;
    }

    try {
        var bounds = getBounds(item, boundsMode);
        return bounds && bounds.length === 4;
    } catch (e) {
        return false;
    }
}

function isLockedOrHidden(item) {
    var current = item;

    while (current) {
        if (hasLockedState(current) || hasHiddenState(current)) {
            return true;
        }
        current = current.parent;
    }

    return false;
}

function hasLockedState(target) {
    try {
        return target.locked === true;
    } catch (e) {
        return false;
    }
}

function hasHiddenState(target) {
    try {
        return target.hidden === true || target.visible === false;
    } catch (e) {
        return false;
    }
}

// =========================================
// 位置を取得・復元 / Get and restore position
// =========================================
function getItemPosition(item) {
    return [item.position[0], item.position[1]];
}

function restoreItemPosition(item, originalPosition) {
    item.position = [originalPosition[0], originalPosition[1]];
}

function copyBounds(bounds) {
    return [bounds[0], bounds[1], bounds[2], bounds[3]];
}


function calculateSwapOffsetsFromBounds(boundsA, boundsB, referenceMode, visualBoundsA, visualBoundsB, relayoutPolicy) {
    if (referenceMode === REFERENCE_MODES.FIXED_EDGES) {
        return calculateFixedEdgeSwapOffsets(boundsA, boundsB, visualBoundsA, visualBoundsB, relayoutPolicy);
    }

    return calculateCenterSwapOffsets(boundsA, boundsB);
}

function calculateCenterSwapOffsets(boundsA, boundsB) {
    var centerA = getCenterFromBounds(boundsA);
    var centerB = getCenterFromBounds(boundsB);
    var offset = getOffset(centerA, centerB);

    return {
        x: offset.x,
        y: offset.y
    };
}

function calculateFixedEdgeSwapOffsets(boundsA, boundsB, visualBoundsA, visualBoundsB, relayoutPolicy) {
    var orderingBoundsA = visualBoundsA || boundsA;
    var orderingBoundsB = visualBoundsB || boundsB;
    var widthA = getWidth(boundsA);
    var widthB = getWidth(boundsB);
    var heightA = getHeight(boundsA);
    var heightB = getHeight(boundsB);
    var isALeftInBase = getCenterFromBounds(boundsA).x <= getCenterFromBounds(boundsB).x;
    var isALeftInVisual = getCenterFromBounds(orderingBoundsA).x <= getCenterFromBounds(orderingBoundsB).x;
    var keepCurrentVisualOrder = (relayoutPolicy === PREVIEW_RELAYOUT_POLICIES.KEEP_CURRENT_VISUAL_ORDER);
    var isALeft = keepCurrentVisualOrder ? isALeftInVisual : isALeftInBase;

    var leftBaseBounds = isALeftInBase ? boundsA : boundsB;
    var rightBaseBounds = isALeftInBase ? boundsB : boundsA;

    var leftSlot = {
        left: leftBaseBounds[0],
        centerY: getBoundsCenterY(leftBaseBounds)
    };
    var rightSlot = {
        right: rightBaseBounds[2],
        centerY: getBoundsCenterY(rightBaseBounds)
    };

    var offsetForA = isALeft ? {
        x: (rightSlot.right - widthA) - boundsA[0],
        y: (rightSlot.centerY + (heightA / 2)) - boundsA[1]
    } : {
        x: leftSlot.left - boundsA[0],
        y: (leftSlot.centerY + (heightA / 2)) - boundsA[1]
    };

    var offsetForB = isALeft ? {
        x: leftSlot.left - boundsB[0],
        y: (leftSlot.centerY + (heightB / 2)) - boundsB[1]
    } : {
        x: (rightSlot.right - widthB) - boundsB[0],
        y: (rightSlot.centerY + (heightB / 2)) - boundsB[1]
    };

    return {
        itemA: offsetForA,
        itemB: offsetForB
    };
}

// =========================================
// 入れ替え用の移動量を適用 / Apply offsets for swapping
// =========================================
function applySwapOffsets(itemA, itemB, offset) {
    if (offset.itemA && offset.itemB) {
        moveItem(itemA, offset.itemA.x, offset.itemA.y);
        moveItem(itemB, offset.itemB.x, offset.itemB.y);
        return;
    }

    moveItem(itemA, offset.x, offset.y);
    moveItem(itemB, -offset.x, -offset.y);
}

// =========================================
// 中心座標を取得 / Get center coordinates
// =========================================
function getCenter(item, boundsMode) {
    return getCenterFromBounds(getBounds(item, boundsMode));
}

function getCenterFromBounds(bounds) {
    return {
        x: (bounds[0] + bounds[2]) / 2,
        y: (bounds[1] + bounds[3]) / 2
    };
}

function getWidth(bounds) {
    return bounds[2] - bounds[0];
}
function areNearlyEqual(a, b) {
    return Math.abs(a - b) < 0.01;
}

function getHeight(bounds) {
    return bounds[1] - bounds[3];
}

function getBoundsCenterY(bounds) {
    return (bounds[1] + bounds[3]) / 2;
}

// =========================================
// bounds を取得 / Get bounds
// =========================================
function getBounds(item, boundsMode) {
    if (boundsMode === BOUNDS_MODES.VISIBLE) {
        return item.visibleBounds;
    }

    if (boundsMode === BOUNDS_MODES.GEOMETRIC) {
        return item.geometricBounds;
    }

    throw new Error(L('errorInvalidBoundsMode') + boundsMode);
}

// =========================================
// 移動量を計算 / Calculate offset
// =========================================
function getOffset(centerA, centerB) {
    return {
        x: centerB.x - centerA.x,
        y: centerB.y - centerA.y
    };
}

// =========================================
// オブジェクトを移動 / Move object
// =========================================
function moveItem(item, dx, dy) {
    item.translate(dx, dy);
}