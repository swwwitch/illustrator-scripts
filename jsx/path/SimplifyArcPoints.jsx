#target illustrator
#targetengine "SimplifyArcPointsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SimplifyArcPoints
Name=SimplifyArcPoints
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-17
Description-BEGIN
選択したパスの直線部分以外を円弧とみなし、90°以下の円弧に置き換えてアンカーポイントを減らします。
ダイアログボックスで結果とアンカーポイント数をプレビューできます。

Treats everything except the straight parts of the selected paths as arcs and rebuilds them with arcs of 90° or less, reducing the number of anchor points.
The dialog previews the result and the anchor point count.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したパスの直線部分以外を円弧とみなし、90°以下の円弧に置き換えてアンカーポイントを減らします。
ダイアログボックスで結果とアンカーポイント数をプレビューできます。

### Overview

Treats everything except the straight parts of the selected paths as arcs and rebuilds them with arcs of 90° or less, reducing the number of anchor points.
The dialog previews the result and the anchor point count.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SimplifyArcPoints";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-17";                   /* 更新日 / last updated */

var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/SimplifyArcPoints.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 方向線のないセグメントを直線とみなす長さ（曲線のセグメントの長さの中央値に対する倍率） / Length ratio over the median curve segment at which a straight segment counts as a line */
    var LINE_LENGTH_RATIO = 2;

    /* 円弧に接する直線のうち、半径に対してこの倍率より短いものは円弧に取り込む / Lines touching an arc shorter than this ratio of its radius join the arc */
    var SHORT_LINE_RATIO = 0.3;

    /* 開いたときにプレビューをオンにする / Start with the preview on */
    var DEFAULT_PREVIEW = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var ANCHOR_COUNT_WIDTH = 200; /* アンカーポイント数の表示幅 / width of the anchor count text */

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
            title: { ja: "円弧のアンカーポイントを単純化", en: "Simplify Arc Points" }
        },
        fieldLabel: {
            anchorPoints: { ja: "アンカーポイント", en: "Anchor points" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noPath: { ja: "パスを選択してください。", en: "Please select a path." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

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
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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
    // パスの読み書き / Path data
    // =========================================

    /**
     * 選択からパスを集める（グループ・複合パスの中も含む）
     * @param {Array} items - 対象のページアイテム
     * @param {PathItem[]} pathItems - 集めたパスの格納先
     * @returns {void}
     */
    function collectPaths(items, pathItems) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "PathItem") {
                pathItems.push(item);
            } else if (item.typename === "GroupItem") {
                collectPaths(item.pageItems, pathItems);
            } else if (item.typename === "CompoundPathItem") {
                collectPaths(item.pathItems, pathItems);
            }
        }
    }

    /**
     * パスのポイントを読み取る
     * @param {PathItem} pathItem - 対象のパス
     * @returns {Object[]} ポイント情報 {a, l, r, t} の配列
     */
    function readPathNodes(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var nodes = [];
        for (var i = 0; i < pathPoints.length; i++) {
            nodes.push({
                a: pathPoints[i].anchor,
                l: pathPoints[i].leftDirection,
                r: pathPoints[i].rightDirection,
                t: pathPoints[i].pointType
            });
        }
        return nodes;
    }

    /**
     * ポイント情報でパスを作り直す
     * @param {PathItem} pathItem - 対象のパス
     * @param {Object[]} nodes - ポイント情報 {a, l, r, t} の配列
     * @returns {void}
     */
    function writePathNodes(pathItem, nodes) {
        var anchors = [];
        for (var i = 0; i < nodes.length; i++) anchors.push(nodes[i].a);
        pathItem.setEntirePath(anchors);
        var pathPoints = pathItem.pathPoints;
        for (var j = 0; j < nodes.length; j++) {
            pathPoints[j].pointType = nodes[j].t;
            pathPoints[j].leftDirection = nodes[j].l;
            pathPoints[j].rightDirection = nodes[j].r;
        }
    }

    /**
     * ポイント情報を複製する
     * @param {Object} node - {a, l, r, t}
     * @returns {Object} 複製
     */
    function copyNode(node) {
        return { a: node.a, l: node.l, r: node.r, t: node.t };
    }

    // =========================================
    // 円弧への置き換え / Arc simplification
    // =========================================

    /**
     * 直線以外の区間を円弧とみなして置き換えたポイント列を求める
     * @param {Object[]} pts - 元のポイント情報
     * @param {boolean} isClosed - クローズパスなら true
     * @returns {Object[]|null} 置き換え後のポイント情報。アンカーが減らなければ null
     */
    function computeSimplifiedNodes(pts, isClosed) {
        var n = pts.length;
        if (n < 3) return null;

        /* 直線部分を除いた全体がひとつの円になるクローズパス / closed path with no straight lines: a full circle */
        var lineFlags = findLineSegments(pts, isClosed);
        if (isClosed && !hasTrue(lineFlags)) {
            var circleNodes = buildFullCircleNodes(pts);
            return (circleNodes && circleNodes.length < n) ? circleNodes : null;
        }

        /* クローズパスは直線のセグメントが最後（先頭へ戻る区間）に来るよう始点をずらし、円弧が始点をまたがないようにする
           For closed paths, rotate so a line closes the path and no arc straddles the start point */
        var work = pts;
        var flags = lineFlags;
        if (isClosed) {
            var firstLine = 0;
            while (!lineFlags[firstLine]) firstLine++;
            work = rotateArray(pts, firstLine + 1);
            flags = rotateArray(lineFlags, firstLine + 1);
        }
        absorbShortLines(work, flags);

        var out = [copyNode(work[0])];
        var s = 0;
        var changed = false;
        while (s < n - 1) {
            if (flags[s]) {
                out.push(copyNode(work[s + 1]));
                s++;
                continue;
            }
            /* 次の直線まで（またはパスの終わりまで）を1つの円弧とみなす / treat everything up to the next line as one arc */
            var e = s + 1;
            while (e < n - 1 && !flags[e]) e++;
            var arc = (e - s >= 2) ? fitArc(work, s, e) : null;
            var nodes = arc ? buildArcNodes(arc, work[s], work[e]) : null;
            if (nodes && nodes.length - 1 < e - s) {
                out[out.length - 1].r = nodes[0].r;
                for (var j = 1; j < nodes.length; j++) out.push(nodes[j]);
                changed = true;
            } else {
                for (var k = s + 1; k <= e; k++) out.push(copyNode(work[k]));
            }
            s = e;
        }
        return changed ? out : null;
    }

    /**
     * 円弧に接する短い直線を円弧の側に取り込む（つなぎ目の半端なセグメント）
     * @param {Object[]} pts - ポイント情報
     * @param {boolean[]} flags - セグメントごとの直線判定（書き換える）
     * @returns {void}
     */
    function absorbShortLines(pts, flags) {
        var lastSeg = pts.length - 2; /* 扱うのはアンカー0〜n-1の間のセグメント / segments between anchors 0..n-1 */
        var changed = true;
        while (changed) {
            changed = false;
            for (var i = 0; i <= lastSeg; i++) {
                if (!flags[i]) continue;
                var len = distance(pts[i].a, pts[i + 1].a);
                var neighbors = [i - 1, i + 1];
                for (var k = 0; k < neighbors.length; k++) {
                    var seg = neighbors[k];
                    if (seg < 0 || seg > lastSeg || flags[seg]) continue;
                    var runStart = seg;
                    var runEnd = seg;
                    while (runStart > 0 && !flags[runStart - 1]) runStart--;
                    while (runEnd < lastSeg && !flags[runEnd + 1]) runEnd++;
                    var arc = fitArc(pts, runStart, runEnd + 1);
                    if (arc && len < arc.r * SHORT_LINE_RATIO) {
                        flags[i] = false;
                        changed = true;
                        break;
                    }
                }
            }
        }
    }

    /**
     * 配列の先頭をずらす
     * @param {Array} items - 元の配列
     * @param {number} offset - 先頭にする位置
     * @returns {Array} ずらした配列
     */
    function rotateArray(items, offset) {
        return items.slice(offset).concat(items.slice(0, offset));
    }

    /**
     * 配列に true が含まれるか
     * @param {boolean[]} flags - 判定する配列
     * @returns {boolean} 含まれれば true
     */
    function hasTrue(flags) {
        for (var i = 0; i < flags.length; i++) {
            if (flags[i]) return true;
        }
        return false;
    }

    /**
     * 直線部分のセグメントを判定する
     * 方向線のない（または弦と同じ向きの）セグメントのうち、曲線のセグメントより十分長いものを直線とみなす。
     * 円弧を細かい直線でなぞったパスでは、弦は短いので円弧の側に入る
     * @param {Object[]} pts - ポイント情報
     * @param {boolean} isClosed - クローズパスなら true
     * @returns {boolean[]} セグメントごとの判定（i は アンカー i → i+1）
     */
    function findLineSegments(pts, isClosed) {
        var n = pts.length;
        var segCount = isClosed ? n : n - 1;
        var lengths = [];
        var straightFlags = [];
        var curveLengths = [];
        for (var i = 0; i < segCount; i++) {
            var p0 = pts[i];
            var p1 = pts[(i + 1) % n];
            var len = distance(p0.a, p1.a);
            var straight = isStraightSegment(p0, p1);
            lengths.push(len);
            straightFlags.push(straight);
            if (!straight) curveLengths.push(len);
        }
        /* 曲線のセグメントがあればその長さ、なければ全体の長さの中央値を基準にする / reference: median of curve lengths, else of all */
        var reference = median(curveLengths.length > 0 ? curveLengths : lengths);
        var flags = [];
        for (var j = 0; j < segCount; j++) {
            flags.push(straightFlags[j] && lengths[j] > reference * LINE_LENGTH_RATIO);
        }
        return flags;
    }

    /**
     * セグメントが直線か（方向線が無いか、弦の上にあるか）
     * @param {Object} p0 - 始点側のポイント
     * @param {Object} p1 - 終点側のポイント
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(p0, p1) {
        var chordLength = distance(p0.a, p1.a);
        if (chordLength < 1e-9) return true;
        return distanceToLine(p0.r, p0.a, p1.a) < chordLength * 1e-4
            && distanceToLine(p1.l, p0.a, p1.a) < chordLength * 1e-4;
    }

    /**
     * 2点間の距離
     * @param {Array} p - [x, y]
     * @param {Array} q - [x, y]
     * @returns {number} 距離
     */
    function distance(p, q) {
        var dx = q[0] - p[0];
        var dy = q[1] - p[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 点から直線（a, b を通る）までの距離
     * @param {Array} p - [x, y]
     * @param {Array} a - 直線上の点
     * @param {Array} b - 直線上の点
     * @returns {number} 距離
     */
    function distanceToLine(p, a, b) {
        var len = distance(a, b);
        return Math.abs((b[0] - a[0]) * (a[1] - p[1]) - (a[0] - p[0]) * (b[1] - a[1])) / len;
    }

    /**
     * 中央値
     * @param {number[]} values - 数値の配列
     * @returns {number} 中央値（空なら 0）
     */
    function median(values) {
        if (values.length === 0) return 0;
        var sorted = values.slice(0).sort(function (x, y) { return x - y; });
        var mid = Math.floor(sorted.length / 2);
        return (sorted.length % 2) ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }

    /**
     * セグメント上の点を求める（3次ベジェ）
     * @param {Object} p0 - 始点側のポイント
     * @param {Object} p1 - 終点側のポイント
     * @param {number} t - 0〜1
     * @returns {Array} [x, y]
     */
    function bezierAt(p0, p1, t) {
        var u = 1 - t;
        var b0 = u * u * u, b1 = 3 * u * u * t, b2 = 3 * u * t * t, b3 = t * t * t;
        return [
            b0 * p0.a[0] + b1 * p0.r[0] + b2 * p1.l[0] + b3 * p1.a[0],
            b0 * p0.a[1] + b1 * p0.r[1] + b2 * p1.l[1] + b3 * p1.a[1]
        ];
    }

    /**
     * アンカーs〜eの区間に円を当てはめ、円弧として求める
     * @param {Object[]} pts - ポイント情報
     * @param {number} s - 開始インデックス
     * @param {number} e - 終了インデックス
     * @returns {Object|null} {cx, cy, r, a0, sweep}。円を求められなければ null
     */
    function fitArc(pts, s, e) {
        var samples = collectArcSamples(pts, s, e, false);
        var circle = fitCircle(samples);
        if (!circle) return null;

        /* 標本を順にたどって、回った角度を足し合わせる / accumulate the angle swept along the samples */
        var sweep = 0;
        var prev = angleOf(samples[0], circle);
        for (var k = 1; k < samples.length; k++) {
            var cur = angleOf(samples[k], circle);
            var d = cur - prev;
            while (d > Math.PI) d -= 2 * Math.PI;
            while (d < -Math.PI) d += 2 * Math.PI;
            sweep += d;
            prev = cur;
        }
        if (Math.abs(sweep) < 1e-6 || Math.abs(sweep) > 2 * Math.PI - 1e-3) return null;

        return {
            cx: circle.cx,
            cy: circle.cy,
            r: circle.r,
            a0: angleOf(pts[s].a, circle),
            sweep: sweep
        };
    }

    /**
     * 区間のアンカーとセグメント途中の点を標本として集める
     * @param {Object[]} pts - ポイント情報
     * @param {number} s - 開始インデックス
     * @param {number} e - 終了インデックス（isClosed なら n で先頭に戻る）
     * @param {boolean} isClosed - 最後のアンカーから先頭へ戻るなら true
     * @returns {Array} [[x, y], ...]
     */
    function collectArcSamples(pts, s, e, isClosed) {
        var n = pts.length;
        var samples = [];
        for (var i = s; i < e; i++) {
            var p0 = pts[i % n];
            var p1 = pts[(i + 1) % n];
            samples.push(p0.a);
            for (var t = 1; t <= 3; t++) samples.push(bezierAt(p0, p1, t / 4));
        }
        if (!isClosed) samples.push(pts[e].a);
        return samples;
    }

    /**
     * クローズパス全体を円とみなし、4つのアンカーの円にする
     * @param {Object[]} pts - ポイント情報
     * @returns {Object[]|null} ポイント情報。円を求められなければ null
     */
    function buildFullCircleNodes(pts) {
        var samples = collectArcSamples(pts, 0, pts.length, true);
        var circle = fitCircle(samples);
        if (!circle) return null;
        /* パスの向き（時計回りか）を面積の符号で決める / direction from the signed area */
        var area = 0;
        for (var i = 0; i < samples.length; i++) {
            var p = samples[i];
            var q = samples[(i + 1) % samples.length];
            area += p[0] * q[1] - q[0] * p[1];
        }
        var step = (area >= 0 ? 1 : -1) * Math.PI / 2;
        var a0 = angleOf(pts[0].a, circle);
        var h = 4 / 3 * Math.tan(step / 4) * circle.r;
        var nodes = [];
        for (var k = 0; k < 4; k++) {
            var ang = a0 + step * k;
            var dx = -Math.sin(ang) * h;
            var dy = Math.cos(ang) * h;
            var anchor = [circle.cx + circle.r * Math.cos(ang), circle.cy + circle.r * Math.sin(ang)];
            nodes.push({
                a: anchor,
                l: [anchor[0] - dx, anchor[1] - dy],
                r: [anchor[0] + dx, anchor[1] + dy],
                t: PointType.SMOOTH
            });
        }
        return nodes;
    }

    /**
     * 中心から見た点の角度
     * @param {Array} p - [x, y]
     * @param {Object} c - {cx, cy}
     * @returns {number} ラジアン
     */
    function angleOf(p, c) {
        return Math.atan2(p[1] - c.cy, p[0] - c.cx);
    }

    /**
     * 点群に円を当てはめる（最小二乗法）
     * @param {Array} samples - [[x, y], ...]
     * @returns {Object|null} {cx, cy, r}
     */
    function fitCircle(samples) {
        var n = samples.length;
        var mx = 0, my = 0;
        for (var i = 0; i < n; i++) {
            mx += samples[i][0];
            my += samples[i][1];
        }
        mx /= n;
        my /= n;

        // 平均を原点にして x²+y²+Dx+Ey+F=0 を解く
        var sxx = 0, sxy = 0, syy = 0, sx = 0, sy = 0, sxz = 0, syz = 0, sz = 0;
        for (var j = 0; j < n; j++) {
            var x = samples[j][0] - mx;
            var y = samples[j][1] - my;
            var z = x * x + y * y;
            sxx += x * x; sxy += x * y; syy += y * y;
            sx += x; sy += y;
            sxz += x * z; syz += y * z; sz += z;
        }
        var m = [[sxx, sxy, sx], [sxy, syy, sy], [sx, sy, n]];
        var b = [-sxz, -syz, -sz];
        var det = det3(m);
        if (Math.abs(det) < 1e-12) return null;

        var sol = [];
        for (var c = 0; c < 3; c++) {
            var mc = [m[0].slice(0), m[1].slice(0), m[2].slice(0)];
            for (var row = 0; row < 3; row++) mc[row][c] = b[row];
            sol.push(det3(mc) / det);
        }
        var cx = -sol[0] / 2;
        var cy = -sol[1] / 2;
        var r2 = cx * cx + cy * cy - sol[2];
        if (r2 <= 0) return null;
        var r = Math.sqrt(r2);
        if (r > 1e5) return null; // ほぼ直線
        return { cx: cx + mx, cy: cy + my, r: r };
    }

    /**
     * 3×3行列の行列式
     * @param {Array} m - 3×3の配列
     * @returns {number} 行列式
     */
    function det3(m) {
        return m[0][0] * (m[1][1] * m[2][2] - m[1][2] * m[2][1])
            - m[0][1] * (m[1][0] * m[2][2] - m[1][2] * m[2][0])
            + m[0][2] * (m[1][0] * m[2][1] - m[1][1] * m[2][0]);
    }

    /**
     * 円弧を90°以下に分けたポイント列を作る
     * 両端は元のアンカー位置を使い、方向線だけ円の接線に合わせる
     * @param {Object} arc - {cx, cy, r, a0, sweep}
     * @param {Object} startNode - 開始ポイント
     * @param {Object} endNode - 終了ポイント
     * @returns {Array} ポイント情報の配列（先頭は開始、末尾は終了）
     */
    function buildArcNodes(arc, startNode, endNode) {
        // 当てはめの誤差で180°がわずかに超えても3分割にならないよう、1°ほどの超過は許す
        var count = Math.max(1, Math.ceil(Math.abs(arc.sweep) / (Math.PI / 2) - 0.02));
        var step = arc.sweep / count;
        var h = 4 / 3 * Math.tan(step / 4) * arc.r;
        var nodes = [];
        for (var i = 0; i <= count; i++) {
            var ang = arc.a0 + step * i;
            var dx = -Math.sin(ang) * h;
            var dy = Math.cos(ang) * h;
            var p = [arc.cx + arc.r * Math.cos(ang), arc.cy + arc.r * Math.sin(ang)];
            if (i === 0) p = startNode.a;
            if (i === count) p = endNode.a;
            nodes.push({
                a: p,
                l: [p[0] - dx, p[1] - dy],
                r: [p[0] + dx, p[1] + dy],
                t: PointType.SMOOTH
            });
        }
        nodes[count].r = endNode.r;
        nodes[count].t = endNode.t;
        return nodes;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー用に作った複製と、一時的に隠した元のアイテム / Preview duplicates and temporarily hidden originals */
    var previewDuplicates = [];
    var previewHiddenItems = [];

    /**
     * 複製に結果を当ててプレビューする（元のアイテムは隠す）
     * @param {Array} sourceItems - 選択していたアイテム
     * @param {Array} simplifiedList - パスごとの置き換え結果（null は変更なし）
     * @returns {void}
     */
    function showPreview(sourceItems, simplifiedList) {
        clearPreview();
        /* 複製は元を隠す前に作る（hidden を引き継ぐため） / duplicate before hiding; duplicates inherit hidden */
        for (var i = 0; i < sourceItems.length; i++) {
            previewDuplicates.push(sourceItems[i].duplicate(sourceItems[i], ElementPlacement.PLACEBEFORE));
        }
        for (var j = 0; j < sourceItems.length; j++) {
            sourceItems[j].hidden = true;
            previewHiddenItems.push(sourceItems[j]);
        }
        var previewPaths = [];
        collectPaths(previewDuplicates, previewPaths);
        if (previewPaths.length === simplifiedList.length) {
            for (var k = 0; k < previewPaths.length; k++) {
                if (simplifiedList[k]) writePathNodes(previewPaths[k], simplifiedList[k]);
            }
        }
        app.redraw();
    }

    /**
     * プレビューの複製を消し、隠したアイテムを戻す
     * @returns {void}
     */
    function clearPreview() {
        for (var i = 0; i < previewDuplicates.length; i++) previewDuplicates[i].remove();
        for (var j = 0; j < previewHiddenItems.length; j++) previewHiddenItems[j].hidden = false;
        var hadPreview = previewDuplicates.length > 0;
        previewDuplicates = [];
        previewHiddenItems = [];
        if (hadPreview) app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 結果をプレビューするダイアログを表示し、確定したらパスを単純化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var sourceItems = [];
        for (var i = 0; i < doc.selection.length; i++) sourceItems.push(doc.selection[i]);
        var targetPaths = [];
        collectPaths(sourceItems, targetPaths);
        if (targetPaths.length === 0) {
            alert(getLabel("alert.noPath"));
            return;
        }

        /* 結果は開く前に一度だけ計算する / compute the result once before opening */
        var simplifiedList = [];
        var beforeCount = 0;
        var afterCount = 0;
        for (var j = 0; j < targetPaths.length; j++) {
            var sourceNodes = readPathNodes(targetPaths[j]);
            var simplified = computeSimplifiedNodes(sourceNodes, targetPaths[j].closed);
            simplifiedList.push(simplified);
            beforeCount += sourceNodes.length;
            afterCount += simplified ? simplified.length : sourceNodes.length;
        }

        var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);
        mainDialog.alignChildren = ["left", "top"];

        var anchorCountText = mainDialog.add("statictext", undefined, labelValueText("fieldLabel.anchorPoints", beforeCount + " → " + afterCount));
        anchorCountText.preferredSize.width = ANCHOR_COUNT_WIDTH;

        var previewCheckbox = mainDialog.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = DEFAULT_PREVIEW;
        var isPreviewOn = DEFAULT_PREVIEW; /* show() 前の value は読み戻せないので控える / value can't be read back before show() */
        previewCheckbox.onClick = function () {
            isPreviewOn = previewCheckbox.value;
            refresh();
        };

        var buttonRow = addButtonRow(mainDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * プレビューのオン・オフを反映する
         * @returns {void}
         */
        function refresh() {
            if (isPreviewOn) {
                showPreview(sourceItems, simplifiedList);
            } else {
                clearPreview();
            }
        }

        refresh();
        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        var dialogResult = mainDialog.show();
        clearPreview();
        if (dialogResult !== 1) return;

        for (var m = 0; m < targetPaths.length; m++) {
            if (simplifiedList[m]) writePathNodes(targetPaths[m], simplifiedList[m]);
        }
    }

    main();

})();
