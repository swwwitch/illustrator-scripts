#target illustrator
#targetengine "SwapObjectsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの位置を入れ替えます。
「中心位置を交換」と「両端の位置を保って交換」を切り替えられ、見た目のサイズを基準にすることもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapObjects.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na534a676fae2

### Overview

Swaps the positions of two selected objects.
You can swap their centers, or keep their outer edges fixed, and optionally work from their visual bounds.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapObjects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapObjects";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapObjects.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapObjects.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na534a676fae2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_USE_VISUAL_BOUNDS      = false;  /* 「見た目のサイズを基準にする」の初期値 / default for visual bounds */
    var DEFAULT_LIVE_PREVIEW_ENABLED   = false;  /* プレビューの初期値 / default for live preview */
    var EQUAL_VALUE_TOLERANCE          = 0.01;   /* 同値とみなす許容差 / tolerance for treating values as equal */

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
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} spacing - 要素間隔。省略時はWINDOW_SPACING
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} spacing - 要素間隔。省略時はPANEL_SPACING
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

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
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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
                ja: "ドキュメントを開いてから実行してください。",
                en: "Please open a document first."
            },
            selectTwoItems: {
                ja: "2つのオブジェクトを選択してください。",
                en: "Please select two objects."
            },
            unsupportedItem: {
                ja: "対応していないオブジェクトが含まれています。",
                en: "The selection contains unsupported objects."
            },
            lockedOrHidden: {
                ja: "ロックまたは非表示のオブジェクトは対象にできません。",
                en: "Locked or hidden objects are not supported."
            }
        }
    };

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * 前提条件を確認してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
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
            return { isValid: false, message: getLabel('alert.selectTwoItems') };
        }

        for (var i = 0; i < targetItems.length; i++) {
            if (!isSwappableItem(targetItems[i])) {
                return { isValid: false, message: getLabel('alert.unsupportedItem') };
            }
            if (isLockedOrHidden(targetItems[i])) {
                return { isValid: false, message: getLabel('alert.lockedOrHidden') };
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

        dialogControls.btnOK.onClick = function () {
            /* プレビュー適用済みならそのまま確定 / Keep the applied preview as the result */
            if (!swapController.isSwapApplied() && swapController.applySwap(getSwapSettings())) {
                app.redraw();
            }
            dialogControls.dialogWindow.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.dialogWindow.close(0);
        };

        prepareDialogWindow(dialogControls.dialogWindow, SCRIPT_NAME);
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
        var dialogWindow = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dialogWindow);

        /* 位置の基準 / Position reference */
        var positionReferencePanel = dialogWindow.add('panel', undefined, getLabel('panel.positionReference'));
        setupPanel(positionReferencePanel, 6);

        var rdoSwapCenters = positionReferencePanel.add('radiobutton', undefined, getLabel('radio.swapCenters'));
        rdoSwapCenters.helpTip = getLabel('tooltip.swapCenters');

        var rdoKeepOuterEdges = positionReferencePanel.add('radiobutton', undefined, getLabel('radio.keepOuterEdges'));
        rdoKeepOuterEdges.helpTip = getLabel('tooltip.keepOuterEdges');

        /* サイズの基準 / Size reference */
        var sizeReferencePanel = dialogWindow.add('panel', undefined, getLabel('panel.sizeReference'));
        setupPanel(sizeReferencePanel, 6);

        var chkUseVisualBounds = sizeReferencePanel.add('checkbox', undefined, getLabel('checkbox.useVisualBounds'));
        chkUseVisualBounds.helpTip = getLabel('tooltip.useVisualBounds');
        chkUseVisualBounds.value = DEFAULT_USE_VISUAL_BOUNDS;

        /* ボタンエリア（左：プレビュー、右：キャンセル / OK）
           プレビューは押しっぱなしの切り替えなので、ボタンではなくチェックボックスにしている
           Button row (left: Preview, right: Cancel / OK).
           Preview is a sticky toggle, so it stays a checkbox rather than a button */
        var buttonRow = addButtonRow(dialogWindow);
        var chkLivePreview = buttonRow.leftGroup.add('checkbox', undefined, getLabel('checkbox.livePreview'));
        chkLivePreview.helpTip = getLabel('tooltip.livePreview');
        chkLivePreview.value = DEFAULT_LIVE_PREVIEW_ENABLED;

        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        return {
            dialogWindow: dialogWindow,
            rdoSwapCenters: rdoSwapCenters,
            rdoKeepOuterEdges: rdoKeepOuterEdges,
            chkUseVisualBounds: chkUseVisualBounds,
            chkLivePreview: chkLivePreview,
            btnCancel: btnCancel,
            btnOK: btnOK
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
     * @param {object} offsetPair - itemA / itemB それぞれの移動量
     * @returns {void}
     */
    function translateItems(targetItems, offsetPair) {
        targetItems[0].translate(offsetPair.itemA.x, offsetPair.itemA.y);
        targetItems[1].translate(offsetPair.itemB.x, offsetPair.itemB.y);
    }

    /**
     * 移動量の符号を反転する
     * @param {object} offsetPair - itemA / itemB それぞれの移動量
     * @returns {object} 符号を反転した移動量
     */
    function negateOffsets(offsetPair) {
        return {
            itemA: { x: -offsetPair.itemA.x, y: -offsetPair.itemA.y },
            itemB: { x: -offsetPair.itemB.x, y: -offsetPair.itemB.y }
        };
    }

    /**
     * 実質移動しない移動量か判定する
     * @param {object} offsetPair - itemA / itemB それぞれの移動量
     * @returns {boolean} 両方とも移動量がゼロとみなせる場合はtrue
     */
    function isZeroOffsetPair(offsetPair) {
        return isNearlyEqual(offsetPair.itemA.x, 0) && isNearlyEqual(offsetPair.itemA.y, 0) &&
            isNearlyEqual(offsetPair.itemB.x, 0) && isNearlyEqual(offsetPair.itemB.y, 0);
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
