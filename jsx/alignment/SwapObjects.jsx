#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの位置を入れ替えます。
「中心位置を交換」と「両端の位置を保って交換」を切り替えられ、見た目のサイズを基準にすることもできます。

詳細は README を参照してください。

### Overview

Swaps the positions of two selected objects.
You can swap their centers, or keep their outer edges fixed, and optionally work from their visual bounds.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapObjects";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapObjects.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapObjects.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na534a676fae2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_USE_VISUAL_BOUNDS      = false;  /* 「見た目のサイズを基準にする」の初期値 / default for visual bounds */
    var DEFAULT_LIVE_PREVIEW_ENABLED   = false;  /* プレビューの初期値 / default for live preview */
    var EQUAL_VALUE_TOLERANCE          = 0.01;   /* 同値とみなす許容差 / tolerance for treating values as equal */

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
        }
    };

    /**
     * 現在のUI言語に応じたラベル文字列を取得する
     * @param {string} category - LABELSのカテゴリ名
     * @param {string} key - カテゴリ内のキー名
     * @returns {string} 対応する文字列
     */
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

    /**
     * ウィンドウに共通のレイアウト設定を適用する
     * @param {Window} win - 対象のウィンドウ
     * @param {number} spacing - 要素間隔。省略時はWINDOW_SPACING
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} panel - 対象のパネル
     * @param {number} spacing - 要素間隔。省略時はPANEL_SPACING
     * @returns {void}
     */
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

    /**
     * 前提条件を確認してダイアログを開く
     * @returns {void}
     */
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

    /**
     * 選択が「2つの入れ替え可能なオブジェクト」か判定する
     * @param {PageItem[]} targetItems - 判定する選択内容
     * @returns {object} isValid（boolean）とmessage（string）を持つ判定結果
     */
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

    /**
     * 移動でき、両方のboundsを取得できるオブジェクトか判定する
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @returns {boolean} 入れ替えの対象にできる場合はtrue
     */
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

    /**
     * boundsが4要素の配列か判定する
     * @param {number[]} bounds - 判定する境界値
     * @returns {boolean} 4要素の配列の場合はtrue
     */
    function isBoundsArray(bounds) {
        return !!bounds && bounds.length === 4;
    }

    /**
     * 自身または祖先がロック・非表示か判定する
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @returns {boolean} ロックまたは非表示の場合はtrue
     */
    function isLockedOrHidden(targetItem) {
        /* ドキュメントより上は見ない（親子関係が循環しても抜けられるようにする）
           Stop at the document so a cyclic parent chain cannot loop forever */
        for (var ancestor = targetItem; ancestor && ancestor.typename !== 'Document'; ancestor = ancestor.parent) {
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

    /**
     * ダイアログを組み立て、プレビューと確定処理を接続する
     * @param {PageItem[]} targetItems - 入れ替える2つのオブジェクト
     * @returns {void}
     */
    function showSwapOptionsDialog(targetItems) {
        var dialogControls = buildSwapDialog();
        var swapController = createSwapController(targetItems);

        /**
         * 「サイズの基準」の状態からbounds modeを得る
         * @returns {string} BOUNDS_MODESのいずれか
         */
        function getBoundsMode() {
            return dialogControls.chkUseVisualBounds.value ? BOUNDS_MODES.VISUAL : BOUNDS_MODES.GEOMETRIC;
        }

        /**
         * ダイアログの現在値を設定オブジェクトにまとめる
         * @returns {object} boundsModeとreferenceModeを持つ設定
         */
        function getSwapSettings() {
            return {
                boundsMode: getBoundsMode(),
                referenceMode: dialogControls.rdoKeepOuterEdges.value ? REFERENCE_MODES.KEEP_OUTER_EDGES : REFERENCE_MODES.SWAP_CENTERS
            };
        }

        /* 初期選択も入れ替えと同じ bounds で判定する / Pick the initial mode from the same bounds as the swap */
        var initialReferenceMode = getInitialReferenceMode(
            swapController.getOriginalBoundsPair(getBoundsMode())
        );
        dialogControls.rdoSwapCenters.value = (initialReferenceMode === REFERENCE_MODES.SWAP_CENTERS);
        dialogControls.rdoKeepOuterEdges.value = (initialReferenceMode === REFERENCE_MODES.KEEP_OUTER_EDGES);

        /**
         * プレビューの適用・解除を切り替えて再描画する
         * @returns {void}
         */
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
            if (!swapController.isSwapApplied() && swapController.applySwap(getSwapSettings())) {
                app.redraw();
            }
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

    /**
     * ダイアログの各コントロールを生成して返す
     * @returns {object} 生成したコントロールをまとめたオブジェクト
     */
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

    /**
     * 初期選択する位置の基準を決める（横並びで幅が違うときだけ両端維持）
     * @param {number[][]} boundsPair - 2つのオブジェクトのbounds
     * @returns {string} REFERENCE_MODESのいずれか
     */
    function getInitialReferenceMode(boundsPair) {
        /* 中心Xが同じ＝縦並び。両端維持は横並び前提なので中心交換にする
           Equal center X means a vertical stack, and keep-edges assumes a horizontal row */
        if (isNearlyEqual(getBoundsCenter(boundsPair[0]).x, getBoundsCenter(boundsPair[1]).x)) {
            return REFERENCE_MODES.SWAP_CENTERS;
        }

        var widthA = getBoundsWidth(boundsPair[0]);
        var widthB = getBoundsWidth(boundsPair[1]);

        return isNearlyEqual(widthA, widthB) ? REFERENCE_MODES.SWAP_CENTERS : REFERENCE_MODES.KEEP_OUTER_EDGES;
    }

    /**
     * 誤差を許容して2つの数値を同値とみなすか判定する
     * @param {number} valueA - 比較する値
     * @param {number} valueB - 比較する値
     * @returns {boolean} 同値とみなせる場合はtrue
     */
    function isNearlyEqual(valueA, valueB) {
        return Math.abs(valueA - valueB) < EQUAL_VALUE_TOLERANCE;
    }

    // =========================================
    // 入れ替えの適用と復元 / Apply and restore the swap
    // =========================================

    /**
     * 元の状態を保持し、入れ替えの適用・復元を行うオブジェクトを作る
     * @param {PageItem[]} targetItems - 入れ替える2つのオブジェクト
     * @returns {object} 適用・復元のメソッドを持つコントローラー
     */
    function createSwapController(targetItems) {
        /* 元の bounds を先にスナップショット。適用は必ず元位置からやり直す
           Snapshot the original bounds; every apply starts over from them */
        var originalBoundsByMode = {};
        originalBoundsByMode[BOUNDS_MODES.GEOMETRIC] = snapshotBoundsPair(targetItems, BOUNDS_MODES.GEOMETRIC);
        originalBoundsByMode[BOUNDS_MODES.VISUAL] = snapshotBoundsPair(targetItems, BOUNDS_MODES.VISUAL);

        /* 適用中の移動量。未適用なら null / Offsets currently applied, or null when nothing is applied */
        var appliedOffsets = null;

        /**
         * 掛けた移動量を逆向きに打ち消して元の位置に戻す
         * @returns {boolean} 実際に動かした場合はtrue
         */
        function restoreOriginalPositions() {
            if (!appliedOffsets) {
                return false;
            }

            /* position 代入で戻すとパターン塗りやグラデーションが取り残されるため translate で戻す
               Assigning position would leave pattern fills and gradients behind, so translate back */
            translateItems(targetItems, negateOffsets(appliedOffsets));
            appliedOffsets = null;
            return true;
        }

        return {
            /**
             * 入れ替えが適用中か返す
             * @returns {boolean} 適用中の場合はtrue
             */
            isSwapApplied: function () {
                return !!appliedOffsets;
            },
            /**
             * 指定モードの元のboundsを返す
             * @param {string} boundsMode - BOUNDS_MODESのいずれか
             * @returns {number[][]} 2つのオブジェクトのbounds
             */
            getOriginalBoundsPair: function (boundsMode) {
                return originalBoundsByMode[boundsMode];
            },
            restoreOriginalPositions: restoreOriginalPositions,
            /**
             * 設定に従って入れ替えを適用する
             * @param {object} swapSettings - boundsModeとreferenceModeを持つ設定
             * @returns {boolean} 実際に動かした場合はtrue
             */
            applySwap: function (swapSettings) {
                var didRestore = restoreOriginalPositions();

                var boundsPair = originalBoundsByMode[swapSettings.boundsMode];
                var swapOffsets = calculateSwapOffsets(boundsPair[0], boundsPair[1], swapSettings.referenceMode);

                /* 実質動かないなら触らない（同寸・同位置のとき）/ Skip when there is nothing to move */
                if (isZeroOffsetPair(swapOffsets)) {
                    return didRestore;
                }

                translateItems(targetItems, swapOffsets);
                appliedOffsets = swapOffsets;
                return true;
            }
        };
    }

    /**
     * 2つのオブジェクトをまとめて移動する
     * @param {PageItem[]} targetItems - 移動する2つのオブジェクト
     * @param {object} offsets - itemA / itemB それぞれの移動量
     * @returns {void}
     */
    function translateItems(targetItems, offsets) {
        targetItems[0].translate(offsets.itemA.x, offsets.itemA.y);
        targetItems[1].translate(offsets.itemB.x, offsets.itemB.y);
    }

    /**
     * 移動量の符号を反転する
     * @param {object} offsets - itemA / itemB それぞれの移動量
     * @returns {object} 符号を反転した移動量
     */
    function negateOffsets(offsets) {
        return {
            itemA: { x: -offsets.itemA.x, y: -offsets.itemA.y },
            itemB: { x: -offsets.itemB.x, y: -offsets.itemB.y }
        };
    }

    /**
     * 実質移動しない移動量か判定する
     * @param {object} offsets - itemA / itemB それぞれの移動量
     * @returns {boolean} 両方とも移動量がゼロとみなせる場合はtrue
     */
    function isZeroOffsetPair(offsets) {
        return isNearlyEqual(offsets.itemA.x, 0) && isNearlyEqual(offsets.itemA.y, 0) &&
            isNearlyEqual(offsets.itemB.x, 0) && isNearlyEqual(offsets.itemB.y, 0);
    }

    /**
     * 2つのオブジェクトのboundsを複製して控える
     * @param {PageItem[]} targetItems - 対象の2つのオブジェクト
     * @param {string} boundsMode - BOUNDS_MODESのいずれか
     * @returns {number[][]} 複製した2つのbounds
     */
    function snapshotBoundsPair(targetItems, boundsMode) {
        return [
            copyBounds(getBoundsByMode(targetItems[0], boundsMode)),
            copyBounds(getBoundsByMode(targetItems[1], boundsMode))
        ];
    }

    /**
     * bounds配列を複製する
     * @param {number[]} bounds - 複製する境界値
     * @returns {number[]} 複製した境界値
     */
    function copyBounds(bounds) {
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    /**
     * 指定モードのboundsを取得する
     * @param {PageItem} targetItem - 対象オブジェクト
     * @param {string} boundsMode - BOUNDS_MODESのいずれか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getBoundsByMode(targetItem, boundsMode) {
        return (boundsMode === BOUNDS_MODES.VISUAL) ? targetItem.visibleBounds : targetItem.geometricBounds;
    }

    // =========================================
    // 移動量の計算 / Calculate offsets
    // =========================================

    /**
     * 位置の基準に応じた移動量を返す
     * @param {number[]} boundsA - 一方のbounds
     * @param {number[]} boundsB - もう一方のbounds
     * @param {string} referenceMode - REFERENCE_MODESのいずれか
     * @returns {object} itemA / itemB それぞれの移動量
     */
    function calculateSwapOffsets(boundsA, boundsB, referenceMode) {
        if (referenceMode === REFERENCE_MODES.KEEP_OUTER_EDGES) {
            return calculateKeepOuterEdgesOffsets(boundsA, boundsB);
        }
        return calculateSwapCentersOffsets(boundsA, boundsB);
    }

    /**
     * 中心位置を交換する移動量を求める
     * @param {number[]} boundsA - 一方のbounds
     * @param {number[]} boundsB - もう一方のbounds
     * @returns {object} itemA / itemB それぞれの移動量
     */
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

    /**
     * 左端・右端を保ったまま入れ替える移動量を求める
     * @param {number[]} boundsA - 一方のbounds
     * @param {number[]} boundsB - もう一方のbounds
     * @returns {object} itemA / itemB それぞれの移動量
     */
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

    /**
     * 移動先となる左右のスロットを作る
     * @param {number[]} leftBounds - 左側にあるオブジェクトのbounds
     * @param {number[]} rightBounds - 右側にあるオブジェクトのbounds
     * @returns {object} left / right のスロット
     */
    function createOuterEdgeSlots(leftBounds, rightBounds) {
        return {
            left: { edgeX: leftBounds[0], isRightEdge: false, centerY: getBoundsCenter(leftBounds).y },
            right: { edgeX: rightBounds[2], isRightEdge: true, centerY: getBoundsCenter(rightBounds).y }
        };
    }

    /**
     * スロットに収めるための移動量を求める
     * @param {number[]} bounds - 移動するオブジェクトのbounds
     * @param {object} targetSlot - 移動先のスロット
     * @returns {object} x / y の移動量
     */
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

    /**
     * boundsの中心座標を返す
     * @param {number[]} bounds - 対象の境界値
     * @returns {object} x / y を持つ中心座標
     */
    function getBoundsCenter(bounds) {
        return {
            x: (bounds[0] + bounds[2]) / 2,
            y: (bounds[1] + bounds[3]) / 2
        };
    }

    /**
     * boundsの幅を返す
     * @param {number[]} bounds - 対象の境界値
     * @returns {number} 幅
     */
    function getBoundsWidth(bounds) {
        return bounds[2] - bounds[0];
    }

    /**
     * boundsの高さを返す（Illustratorは上方向が正のY）
     * @param {number[]} bounds - 対象の境界値
     * @returns {number} 高さ
     */
    function getBoundsHeight(bounds) {
        return bounds[1] - bounds[3];
    }

    main();

})();
