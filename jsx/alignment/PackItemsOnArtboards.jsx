#target illustrator
#targetengine "PackItemsOnArtboardsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PackItemsOnArtboards
Name=オブジェクトをアートボードに詰めて配置
Author=m1b
Target-App=Illustrator
Version=1.2.0
Release-Date=2024-10-22
Description-BEGIN
選択したオブジェクトを、すき間を詰めてアートボードに配置します。
間隔・余白・回転の可否・試行回数はダイアログボックスで指定できます。

Packs the selected items tightly onto the artboards.
Spacing, margin, rotation and the number of attempts are set in a dialog.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトを、すき間を詰めてアートボードに配置します。
間隔・余白・回転の可否・試行回数はダイアログボックスで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PackItemsOnArtboards.md

### Overview

Packs the selected items tightly onto the artboards.
Spacing, margin, rotation and the number of attempts are set in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PackItemsOnArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PackItemsOnArtboards";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "m1b";                          /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2024-10-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-07";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "オブジェクトをアートボードに詰めて配置"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PackItemsOnArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PackItemsOnArtboards.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/PackItemsOnArtboards.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 元のスクリプト: Bin Packing--Artboards.js（InDesign 版は Bin Packing--Pages.js）
 * @author m1b
 * @version 2024-10-22
 * @discussion https://community.adobe.com/t5/illustrator-discussions/how-to-organize-multiple-different-objects-on-one-sheet-with-a-defined-gap-inbetween-them/m-p/12475475#M295934
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var LENGTH_FIELD_UNIT = " mm"; /* 間隔・余白の欄の単位（mm / cm / in / pt）/ unit of the spacing fields */

    var DEFAULT_PACK_SETTINGS = {
        padding: "1 mm",                  /* オブジェクトの間隔（マイナスで重なる）/ space between items (negative values overlap) */
        margin: "5 mm",                   /* アートボードの余白（マイナスも可）/ space around the artboard edges */
        allow90DegreeRotation: false,     /* 入りきらないときに90°回転してよいか / may items rotate 90° to fit */
        allowAnyRotation: false,          /* 配置の前に、外接する長方形が最小になる角度へ回転してよいか / may items first rotate to their smallest bounding rectangle */
        bestFitBy: "count",               /* 優先する基準（"count" は個数、"area" は面積）/ prefer item count or packed area */
        maxAttemptCount: undefined,       /* 試行回数の上限（undefined なら個数から計算）/ maximum attempts (undefined to calculate) */
        tryHarder: false,                 /* すべて配置できても上限まで試すか / keep trying after everything fits */
        doNotSort: false,                 /* 重ね順のまま1回だけ配置するか / pack once in stacking order */
        showUI: true,                     /* ダイアログボックスを表示するか / show the dialog */
        showResults: false,               /* 配置のあとに結果を表示するか / show the results afterwards */
        keepRemainingItemsSelected: true, /* 配置できなかったオブジェクトだけを選択状態にするか / leave only the unpacked items selected */

        /* デバッグ用 / debugging options */
        showBlockBounds: false,
        showOnlyBlockBounds: false,
        debugBinBounds: false,
        forceRotate: false
    };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_LABEL_WIDTH    = 130;         /* 項目名の幅 / field label width */
    var FIELD_CHARACTERS     = 7;           /* 数値欄の文字数 / characters in the numeric fields */
    var RESULT_TEXT_SIZE     = [250, 230];  /* 結果の文字欄の大きさ / size of the results text */
    var PROGRESS_BAR_SIZE    = [388, 12];   /* 進捗バーの大きさ / size of the progress bars */
    var PROGRESS_COUNT_WIDTH = 100;         /* 進捗の件数の幅 / width of the progress count */

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

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
            title: { ja: "アートボードに詰めて配置", en: "Pack Items onto Artboards" },
            progress: { ja: "配置中", en: "Packing Items" },
            results: { ja: "配置結果", en: "Packing Results" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            attempts: { ja: "試行", en: "Attempts" },
            rotation: { ja: "回転", en: "Rotation" }
        },
        fieldLabel: {
            padding: { ja: "オブジェクトの間隔", en: "Space between items" },
            margin: { ja: "アートボードの余白", en: "Artboard margin" },
            maxAttempts: { ja: "最大試行回数", en: "Max attempts" },
            bestFit: { ja: "優先", en: "Maximize" }
        },
        dropdown: {
            count: { ja: "配置した個数", en: "Items packed" },
            area: { ja: "配置した面積", en: "Area packed" }
        },
        checkbox: {
            allow90Rotation: { ja: "90°回転を許可", en: "Allow 90° rotation" },
            allowAnyRotation: { ja: "任意の角度の回転を許可", en: "Allow any rotation" },
            tryHarder: { ja: "最大試行回数まで試す", en: "Use all attempts" },
            doNotSort: { ja: "重ね順のまま配置", en: "Keep stacking order" },
            showResults: { ja: "結果を表示", en: "Show results" }
        },
        button: {
            random: { ja: "ランダムに試す", en: "Try Random" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            pack: { ja: "実行", en: "Pack" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            padding: {
                ja: "オブジェクトどうしの間隔。マイナスの値で重ねて配置します",
                en: "Space between items. Negative values make them overlap"
            },
            margin: {
                ja: "アートボードの端からの余白。マイナスの値も指定できます",
                en: "Space inside the artboard edges. Negative values are allowed"
            },
            maxAttempts: {
                ja: "配置を試す回数の上限。すべて配置できたら途中で終わります",
                en: "Maximum number of attempts. Stops early once every item fits"
            },
            bestFit: {
                ja: "試行の結果を比べるときに、多いほうを優先する値",
                en: "Which value counts most when comparing attempts"
            },
            allow90Rotation: {
                ja: "入りきらないオブジェクトを90°回転して配置します",
                en: "Rotates an item by 90° when it does not fit otherwise"
            },
            allowAnyRotation: {
                ja: "配置の前に、外接する長方形が最小になる角度へ回転します",
                en: "Rotates each item to its smallest bounding rectangle before packing"
            },
            tryHarder: {
                ja: "すべて配置できても、最大試行回数までよりよい結果を探します",
                en: "Keeps looking for a better result after every item fits"
            },
            doNotSort: {
                ja: "大きさで並べ替えず、重ね順のまま1回だけ配置します",
                en: "Packs once in stacking order, without sorting by size"
            },
            showResults: {
                ja: "配置のあとに、試行の内容とアートボードごとの個数を表示します",
                en: "Shows the attempt used and the count per artboard afterwards"
            },
            random: {
                ja: "ダイアログボックスを閉じずに、ランダムな順で1回配置します。［キャンセル］で閉じても配置は残ります",
                en: "Packs once in random order without closing the dialog. The result stays even if you click Cancel"
            },
            /* ステップボタン用 / for the stepper */
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        message: {
            intro: { ja: "{items} 個を {artboards} 枚のアートボードに配置します", en: "Packing {items} item(s) onto {artboards} artboard(s)" }
        },
        progress: {
            packed: { ja: "配置済み", en: "Items packed" },
            attempt: { ja: "試行回数", en: "Attempt" },
            bestAttempt: { ja: "{attempt} 回目の試行で {artboards} 枚に配置", en: "Attempt {attempt}: packed onto {artboards} artboard(s)" }
        },
        result: {
            success: { ja: "{count} 個のオブジェクトを配置しました。", en: "Packed {count} item(s)." },
            failure: { ja: "{count} 個のオブジェクトを配置できませんでした。", en: "{count} item(s) could not be packed." },
            attemptUsed: { ja: "採用した試行", en: "Attempt used" },
            sortMethod: { ja: "並べ替え", en: "Sort method" },
            score: { ja: "スコア", en: "Score" },
            packedOnArtboard: { ja: "アートボード {artboard} に {count} 個を配置しました。", en: "Packed {count} item(s) onto artboard {artboard}." }
        },
        sortMethod: {
            area: { ja: "面積", en: "Area" },
            largestSide: { ja: "長辺", en: "Largest dimension" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            interleave: { ja: "大小交互", en: "Alternating large and small" },
            random: { ja: "ランダム", en: "Random shuffle" },
            none: { ja: "なし", en: "None" }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してから実行してください。", en: "Please select some items and try again." }
        }
    };

    // =========================================
    // 配置の処理 / Packing
    // =========================================

    /**
     * アートボード1枚分の入れ物
     * @typedef {Object} PackingBin
     * @property {number[]} bounds - 余白を除いた範囲 [上, 左, 下, 右]（Y は下向きが正）
     * @property {number} width - 詰める幅（右端の間隔の分を足した値）
     * @property {number} height - 詰める高さ（下端の間隔の分を足した値）
     */

    /**
     * 配置を実行し、進捗ウィンドウを閉じて結果を表示する
     * @param {Object} packSettings - 配置の設定
     * @param {boolean} isRandomAttempt - ランダムな順で1回だけ試すなら true
     * @returns {void}
     */
    function packAndReport(packSettings, isRandomAttempt) {
        var bestAttempt = packItems(packSettings, isRandomAttempt);
        if (packSettings.progressWindow) packSettings.progressWindow.close();
        if (bestAttempt && packSettings.showResults) showResultsDialog(bestAttempt);
    }

    /**
     * 選択したオブジェクトをアートボードに配置する
     * @param {Object} packSettings - 配置の設定（padding と margin はここで pt の数値に置き換わる）
     * @param {boolean} isRandomAttempt - ランダムな順で1回だけ試すなら true
     * @returns {Attempt|undefined} 採用した試行（入れ物の確認だけのデバッグ時は undefined）
     */
    function packItems(packSettings, isRandomAttempt) {
        /* Block がこの2つを pt の数値として読む / Block reads these two as points */
        packSettings.padding = parseLengthField(packSettings.padding) || 0;
        packSettings.margin = parseLengthField(packSettings.margin) || 0;

        var targetDoc = packSettings.doc;
        var targetItems = packSettings.items;
        var progressWindow = packSettings.progressWindow;
        var bins = createBins(targetDoc, packSettings.padding, packSettings.margin);

        if (packSettings.debugBinBounds) {
            drawBinBounds(targetDoc, bins);
            return undefined;
        }

        if (progressWindow) progressWindow.setItemsPackedProgress(0, targetItems.length);
        if (packSettings.allowAnyRotation) rotateItemsToSmallestBounds(targetItems);

        var bestAttempt = findBestAttempt(packSettings, bins, isRandomAttempt);

        for (var i = 0; i < bestAttempt.packedBlocks.length; i++) {
            bestAttempt.packedBlocks[i].positionItemOnArtboard(packSettings);
        }
        if (progressWindow) progressWindow.setItemsPackedProgress(bestAttempt.packedBlocks.length, targetItems.length);
        if (packSettings.keepRemainingItemsSelected) selectRemainingItems(targetDoc, bestAttempt);

        return bestAttempt;
    }

    /**
     * ドキュメントのアートボードから、余白を除いた入れ物を作る
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {number} paddingPt - オブジェクトの間隔（pt）
     * @param {number} marginPt - アートボードの余白（pt）
     * @returns {PackingBin[]} 入れ物の配列
     */
    function createBins(targetDoc, paddingPt, marginPt) {
        var bins = [];
        for (var i = 0; i < targetDoc.artboards.length; i++) {
            var binBounds = artboardRectToBinBounds(targetDoc.artboards[i].artboardRect);
            binBounds = [binBounds[0] + marginPt, binBounds[1] + marginPt, binBounds[2] - marginPt, binBounds[3] - marginPt];
            bins.push({
                bounds: binBounds,
                width: binBounds[3] - binBounds[1] + paddingPt,
                height: binBounds[2] - binBounds[0] + paddingPt
            });
        }
        return bins;
    }

    /**
     * アートボードの範囲 [左, 上, 右, 下]（Y は上向きが正）を、入れ物の範囲 [上, 左, 下, 右]（Y は下向きが正）に直す
     * @param {number[]} artboardRect - アートボードの範囲
     * @returns {number[]} 入れ物の範囲
     */
    function artboardRectToBinBounds(artboardRect) {
        return [-artboardRect[1], artboardRect[0], -artboardRect[3], artboardRect[2]];
    }

    /**
     * デバッグ用に、入れ物の範囲を長方形で描く
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PackingBin[]} bins - 入れ物の配列
     * @returns {void}
     */
    function drawBinBounds(targetDoc, bins) {
        for (var i = 0; i < bins.length; i++) {
            drawRectangle(targetDoc, bins[i].bounds, {
                strokeColor: targetDoc.swatches[4],
                fillColor: targetDoc.swatches[0],
                note: "Bin " + i
            });
        }
    }

    /**
     * オブジェクトを、外接する長方形が最小になる角度へ回転する
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {void}
     */
    function rotateItemsToSmallestBounds(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].rotate(-findRotationByMinimalBounds(targetItems[i]));
        }
    }

    /**
     * 配置を繰り返し試して、スコアが最も高い試行を返す
     * @param {Object} packSettings - 配置の設定
     * @param {PackingBin[]} bins - 入れ物の配列
     * @param {boolean} isRandomAttempt - ランダムな順で1回だけ試すなら true
     * @returns {Attempt} 採用した試行
     */
    function findBestAttempt(packSettings, bins, isRandomAttempt) {
        var totalItemCount = packSettings.items.length;
        var maxAttemptCount = isRandomAttempt ? 1 : (packSettings.maxAttemptCount || getMaxAttemptCount(totalItemCount));
        var progressWindow = packSettings.progressWindow;
        var totalItemArea = 0;
        var bestAttempt;

        for (var attemptIndex = 0; attemptIndex < maxAttemptCount; attemptIndex++) {
            if (progressWindow) progressWindow.setAttemptProgress(attemptIndex + 1, maxAttemptCount);

            var attempt = new Attempt(attemptIndex, bins);
            attempt.remainingBlocks = createBlocks(packSettings);
            if (attemptIndex === 0) totalItemArea = sumBlockAreas(attempt.remainingBlocks);

            /* 試行の番号で並べ替えの方法を順に使い、尽きたらランダムに並べる / cycle through the sort methods, then shuffle */
            if (!packSettings.doNotSort) sortBlocks(attempt, isRandomAttempt ? undefined : attemptIndex);

            packAttemptIntoBins(attempt, bins, packSettings, totalItemArea);

            if (!bestAttempt || attempt.score > bestAttempt.score) {
                bestAttempt = attempt;
                if (progressWindow) {
                    progressWindow.setItemsPackedProgress(bestAttempt.packedBlocks.length, totalItemCount);
                    progressWindow.setBestAttempt(bestAttempt.binCount, bestAttempt.index);
                }
            }

            /* 並べ替えないなら結果は毎回同じ / the result never changes without sorting */
            if (packSettings.doNotSort) break;

            /* 並べ替えの方法をひととおり試し、すべて配置できていれば打ち切る / stop once every sort method is tried and everything fits */
            if (!packSettings.tryHarder && attemptIndex >= BLOCK_SORT_METHODS.length && 0 === bestAttempt.remainingBlocks.length) break;
        }

        return bestAttempt;
    }

    /**
     * 位置の情報を持つブロックを、対象のオブジェクトの数だけ作る
     * @param {Object} packSettings - 配置の設定
     * @returns {Block[]} ブロックの配列
     */
    function createBlocks(packSettings) {
        var blocks = [];
        for (var i = 0; i < packSettings.items.length; i++) {
            blocks.push(new Block(packSettings, packSettings.items[i]));
        }
        return blocks;
    }

    /**
     * ブロックの面積の合計を返す
     * @param {Block[]} blocks - ブロックの配列
     * @returns {number} 面積の合計
     */
    function sumBlockAreas(blocks) {
        var totalArea = 0;
        for (var i = 0; i < blocks.length; i++) {
            totalArea += getBlockArea(blocks[i]);
        }
        return totalArea;
    }

    /**
     * 残りのブロックを入れ物へ順に詰め、試行のスコアを計算する
     * @param {Attempt} attempt - 今回の試行
     * @param {PackingBin[]} bins - 入れ物の配列
     * @param {Object} packSettings - 配置の設定
     * @param {number} totalItemArea - すべてのブロックの面積の合計
     * @returns {void}
     */
    function packAttemptIntoBins(attempt, bins, packSettings, totalItemArea) {
        var prefersCount = ("count" === packSettings.bestFitBy);
        var totalItemCount = packSettings.items.length;

        for (var i = 0; i < bins.length; i++) {
            var bin = bins[i];
            var binPacker = new Packer(bin.width, bin.height, packSettings.allow90DegreeRotation);
            var fitResult = binPacker.fit(attempt.remainingBlocks, i);

            attempt.area += fitResult.area;
            attempt.binCount = i + 1;
            attempt.packedBlocks = attempt.packedBlocks.concat(fitResult.packedBlocks);
            attempt.remainingBlocks = fitResult.remainingBlocks;

            var scoreFactor = prefersCount ? totalItemCount / fitResult.count : totalItemArea / fitResult.area;
            attempt.score += ((bin.width * bin.height) / fitResult.area) * scoreFactor;
            attempt.info.push(getLabel("result.packedOnArtboard", { count: fitResult.count, artboard: i + 1 }));

            if (0 === attempt.remainingBlocks.length) break;
        }

        /* 使うアートボードが少ない試行を優先し、配置できなかった分を引く / favour fewer artboards, penalize unpacked items */
        attempt.score += (bins.length - attempt.binCount) * 100;
        attempt.score -= attempt.remainingBlocks.length * 100;
    }

    /**
     * 配置できなかったオブジェクトだけを選択する
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {Attempt} bestAttempt - 採用した試行
     * @returns {void}
     */
    function selectRemainingItems(targetDoc, bestAttempt) {
        var remainingItems = [];
        for (var i = 0; i < bestAttempt.remainingBlocks.length; i++) {
            remainingItems.push(bestAttempt.remainingBlocks[i].item);
        }
        targetDoc.selection = remainingItems;
    }

    /**
     * 試行回数の上限の目安を返す
     * @param {number} itemCount - オブジェクトの数
     * @returns {number} 試行回数の上限
     */
    function getMaxAttemptCount(itemCount) {
        return Math.min(200, 4 + Math.floor(Math.log(itemCount) / Math.log(2) * 5));
    }

    // =========================================
    // ブロックの並べ替え / Block sorting
    // =========================================

    /* 値の大きいブロックを先にする方法。試行の番号で順に使う / Largest value first, used in turn by attempt index */
    var BLOCK_SORT_METHODS = [
        { name: "area", getSortValue: getBlockArea },
        { name: "largestSide", getSortValue: getBlockLargestSide },
        { name: "width", getSortValue: getBlockWidth },
        { name: "height", getSortValue: getBlockHeight },
        { name: "interleave", getSortValue: getBlockArea, isInterleaved: true }
    ];

    /* 並べ替えの文字列キーは、この値から引いて降順にする。整数部が常に11桁になり、面積の最大（約2.7e8）より十分大きい。
       ExtendScript の toFixed() は有効桁が約15桁なので、これより大きくすると小数が丸められて同値になる
       Sort keys subtract from this to sort descending: always 11 integer digits, well above the largest area.
       ExtendScript's toFixed() keeps about 15 significant digits, so a larger offset rounds the decimals away */
    var SORT_KEY_OFFSET = 5e10;
    var SORT_KEY_INDEX_BASE = 1000000; /* 元の順番の桁をそろえるために足す値 / keeps the original index a fixed width */

    /**
     * ブロックの面積を返す
     * @param {Block} block - ブロック
     * @returns {number} 面積
     */
    function getBlockArea(block) {
        return block.w * block.h;
    }

    /**
     * ブロックの長辺の長さを返す
     * @param {Block} block - ブロック
     * @returns {number} 長辺の長さ
     */
    function getBlockLargestSide(block) {
        return Math.max(block.w, block.h);
    }

    /**
     * ブロックの幅を返す
     * @param {Block} block - ブロック
     * @returns {number} 幅
     */
    function getBlockWidth(block) {
        return block.w;
    }

    /**
     * ブロックの高さを返す
     * @param {Block} block - ブロック
     * @returns {number} 高さ
     */
    function getBlockHeight(block) {
        return block.h;
    }

    /**
     * 試行の残りのブロックを並べ替え、使った方法を attempt.sortType に控える
     * @param {Attempt} attempt - 今回の試行
     * @param {number} [sortMethodIndex] - BLOCK_SORT_METHODS の番号（無い番号や省略時はランダム）
     * @returns {void}
     */
    function sortBlocks(attempt, sortMethodIndex) {
        var sortMethod = BLOCK_SORT_METHODS[sortMethodIndex];
        if (!sortMethod) {
            shuffle(attempt.remainingBlocks);
            attempt.sortType = "random";
            return;
        }
        var sortedBlocks = sortBlocksDescending(attempt.remainingBlocks, sortMethod.getSortValue);
        attempt.remainingBlocks = sortMethod.isInterleaved ? interleaveBlocks(sortedBlocks) : sortedBlocks;
        attempt.sortType = sortMethod.name;
    }

    /**
     * ブロックを値の大きい順に並べた配列を返す。
     * ExtendScript の比較関数つき sort() は遅く並びも狂うので、文字列キーを引数なしの sort() で並べる。同じ値は元の順番を保つ
     * @param {Block[]} blocks - ブロックの配列
     * @param {Function} getSortValue - ブロックから比べる値を返す関数
     * @returns {Block[]} 並べ替えた配列
     */
    function sortBlocksDescending(blocks, getSortValue) {
        var sortKeys = [];
        for (var i = 0; i < blocks.length; i++) {
            sortKeys.push((SORT_KEY_OFFSET - getSortValue(blocks[i])).toFixed(3) + "\u0001" + (SORT_KEY_INDEX_BASE + i));
        }
        sortKeys.sort();

        var sortedBlocks = [];
        for (var j = 0; j < sortKeys.length; j++) {
            sortedBlocks.push(blocks[Number(sortKeys[j].split("\u0001")[1]) - SORT_KEY_INDEX_BASE]);
        }
        return sortedBlocks;
    }

    /**
     * 大きい順に並んだブロックを、大きい半分と小さい半分から交互に取った配列にする
     * @param {Block[]} sortedBlocks - 大きい順のブロック
     * @returns {Block[]} 交互に並べた配列
     */
    function interleaveBlocks(sortedBlocks) {
        var largerHalfCount = Math.ceil(sortedBlocks.length / 2);
        var interleavedBlocks = [];
        for (var i = 0; i < largerHalfCount; i++) {
            interleavedBlocks.push(sortedBlocks[i]);
            if (i + largerHalfCount < sortedBlocks.length) interleavedBlocks.push(sortedBlocks[i + largerHalfCount]);
        }
        return interleavedBlocks;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * 表示ラベルから単位を探す（「H」は「Q」と同じ単位）
     * @param {string} label - 単位の表示ラベル（"mm" など）
     * @returns {{label: string, pointsPerUnit: number}|null} UNITS の要素（見つからなければ null）
     */
    function findUnitByLabel(label) {
        var unitLabel = (label === "H") ? "Q" : label;
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].label === unitLabel) return UNITS[i];
        }
        return null;
    }

    // =========================================
    // 配置の計算 / Packing engine
    // =========================================

    /* Packer.js と packer-blocks.js のうち、Illustrator で使う部分を取り込んで書き直したもの
       The Illustrator parts of Packer.js and packer-blocks.js, bundled and rewritten here

       Packer: trentium（m1b が ExtendScript 向けに改変）/ by trentium, adapted for ExtendScript by m1b
       https://stackoverflow.com/questions/56642111/bin-packing-js-implementation-using-box-rotation-for-best-fit

       MIT License

       Copyright (c) 2024 mark1bean

       Permission is hereby granted, free of charge, to any person obtaining a copy
       of this software and associated documentation files (the "Software"), to deal
       in the Software without restriction, including without limitation the rights
       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
       copies of the Software, and to permit persons to whom the Software is
       furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all
       copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
       SOFTWARE. */

    var ROTATION_CONVERGENCE_THRESHOLD = 0.001; /* 最小の外接長方形を探す角度の刻みの下限（°）/ smallest angle step when searching for the minimal bounds */
    var ROTATION_START_STEP = 45;               /* 最初の角度の刻み（°）/ first angle step */

    /**
     * 配置の試行1回分の記録
     * @constructor
     * @param {number} attemptIndex - 試行の番号（0始まり）
     * @param {PackingBin[]} bins - 入れ物の配列
     */
    function Attempt(attemptIndex, bins) {
        this.index = attemptIndex;
        this.area = 0;
        this.binCount = bins.length;
        this.score = 0;
        this.info = [];
        this.packedBlocks = [];
        this.remainingBlocks = [];
    }

    /**
     * 配置中のオブジェクト1つの大きさと位置を控える
     * @constructor
     * @param {Object} packSettings - 配置の設定（padding は pt の数値）
     * @param {PageItem} item - 対象のオブジェクト
     */
    function Block(packSettings, item) {
        var bounds = getItemBounds(item);

        this.item = item;
        this.isRotated = false;
        this.binIndex = undefined; /* 配置が決まると Packer が入れる / set by Packer once packed */

        this.w = bounds[2] - bounds[0] + packSettings.padding;
        this.h = bounds[1] - bounds[3] + packSettings.padding;
        this.dx = item.left - bounds[0];
        this.dy = item.top - bounds[1];

        this.dimensions = { w: this.w, h: this.h, dx: this.dx, dy: this.dy };
        this.rotatedDimensions = { w: this.h, h: this.w, dx: -this.dy, dy: item.visibleBounds[2] - bounds[2] };

        if (true === packSettings.forceRotate) this.rotate();
    }

    /**
     * 0°と90°の向きを切り替える
     * @returns {void}
     */
    Block.prototype.rotate = function () {
        if (this.w == this.h) return;

        this.isRotated = !this.isRotated;
        var dimensions = this.isRotated ? this.rotatedDimensions : this.dimensions;
        this.w = dimensions.w;
        this.h = dimensions.h;
        this.dx = dimensions.dx;
        this.dy = dimensions.dy;
    };

    /**
     * オブジェクトをアートボード上の決まった位置へ動かす
     * @param {Object} packSettings - 配置の設定（margin は pt の数値）
     * @returns {void}
     */
    Block.prototype.positionItemOnArtboard = function (packSettings) {
        if (!this.packed) return;
        if (this.isRotated) this.item.rotate(90);

        var artboardRect = packSettings.doc.artboards[this.binIndex].artboardRect;
        var blockLeft = this.x0 + artboardRect[0] + packSettings.margin;
        var blockTop = -(this.y0 - artboardRect[1] + packSettings.margin);

        if (packSettings.showBlockBounds) {
            var blockRight = this.x1 + artboardRect[0] + packSettings.margin;
            var blockBottom = -(this.y1 - artboardRect[1] + packSettings.margin);
            drawRectangle(this.item.parent, [blockLeft, blockTop, blockRight, blockBottom]);
        }
        if (packSettings.showOnlyBlockBounds) return;

        this.item.left = blockLeft + this.dx;
        this.item.top = blockTop + this.dy;
    };

    /**
     * 入れ物1つ分の空き領域を管理し、ブロックを詰める。
     * 空き領域（heap）は重なりを許した長方形の配列で、ブロックを置くたびに分割と統合をする
     * @constructor
     * @param {number} binWidth - 入れ物の幅
     * @param {number} binHeight - 入れ物の高さ
     * @param {boolean} allow90DegreeRotation - 入らないときに90°回転してよいか
     */
    function Packer(binWidth, binHeight, allow90DegreeRotation) {
        this.binWidth = binWidth;
        this.binHeight = binHeight;
        this.allow90DegreeRotation = allow90DegreeRotation;
        this.heap = [];
    }

    /**
     * ブロックを順に入れ物へ詰める
     * @param {Block[]} blocks - 詰めるブロック
     * @param {number} binIndex - 入れ物の番号
     * @returns {{count: number, area: number, packedBlocks: Block[], remainingBlocks: Block[]}} 詰めた結果
     */
    Packer.prototype.fit = function (blocks, binIndex) {
        this.heap = [{ x0: 0, y0: 0, x1: this.binWidth, y1: this.binHeight }];

        var packedArea = 0;
        var packedBlocks = [];
        var remainingBlocks = [];

        for (var i = 0; i < blocks.length; i++) {
            var block = blocks[i];
            var isPlaced = this.findInHeap(block);

            if (!isPlaced && this.allow90DegreeRotation) {
                block.rotate();
                isPlaced = this.findInHeap(block);
            }

            if (isPlaced) {
                this.adjustHeap(block);
                block.binIndex = binIndex;
                packedBlocks.push(block);
                packedArea += block.w * block.h;
            } else {
                remainingBlocks.push(block);
            }
        }

        return { count: packedBlocks.length, area: packedArea, packedBlocks: packedBlocks, remainingBlocks: remainingBlocks };
    };

    /**
     * ブロックが入る空き領域を探し、見つかればその左上に置く
     * @param {Block} block - 置くブロック
     * @returns {boolean} 置けたら true
     */
    Packer.prototype.findInHeap = function (block) {
        for (var i = 0; i < this.heap.length; i++) {
            var heapBlock = this.heap[i];
            if (heapBlock && block.w <= heapBlock.x1 - heapBlock.x0 && block.h <= heapBlock.y1 - heapBlock.y0) {
                block.x0 = heapBlock.x0;
                block.y0 = heapBlock.y0;
                block.x1 = heapBlock.x0 + block.w;
                block.y1 = heapBlock.y0 + block.h;
                block.packed = true;
                return true;
            }
        }
        return false;
    };

    /**
     * 置いたブロックと重なる空き領域を、重ならない上・右・下・左の領域に分割する
     * @param {Block} block - 置いたブロック
     * @returns {void}
     */
    Packer.prototype.adjustHeap = function (block) {
        var heapCount = this.heap.length; /* 分割で増えた分は対象にしない / newly added regions are not revisited */

        for (var i = 0; i < heapCount; i++) {
            var heapBlock = this.heap[i];
            var overlap = getBlockOverlap(heapBlock, block);
            if (!overlap) continue;

            if (overlap.y1 !== heapBlock.y1) this.heap.push({ x0: heapBlock.x0, y0: overlap.y1, x1: heapBlock.x1, y1: heapBlock.y1 });
            if (overlap.x1 !== heapBlock.x1) this.heap.push({ x0: overlap.x1, y0: heapBlock.y0, x1: heapBlock.x1, y1: heapBlock.y1 });
            if (heapBlock.y0 !== overlap.y0) this.heap.push({ x0: heapBlock.x0, y0: heapBlock.y0, x1: heapBlock.x1, y1: overlap.y0 });
            if (heapBlock.x0 !== overlap.x0) this.heap.push({ x0: heapBlock.x0, y0: heapBlock.y0, x1: overlap.x0, y1: heapBlock.y1 });

            this.heap[i] = null;
        }

        this.unionAll();
    };

    /**
     * 空き領域どうしを見比べて、隣り合う領域を広げ、ほかに含まれる領域を除く
     * @returns {void}
     */
    Packer.prototype.unionAll = function () {
        var heap = this.heap;

        for (var i = 0; i < heap.length; i++) {
            for (var j = 0; j < heap.length; j++) {
                if (i === j) continue;

                expandOverlappingBlocks(heap[i], heap[j]);
                if (!heap[i] || !heap[j]) continue;

                if (blockContains(heap[j], heap[i])) {
                    heap[i] = null;
                } else if (blockContains(heap[i], heap[j])) {
                    heap[j] = null;
                }
            }
        }

        var remainingHeap = [];
        for (var k = 0; k < heap.length; k++) {
            if (heap[k]) remainingHeap.push(heap[k]);
        }
        this.heap = remainingHeap;
    };

    /**
     * 2つの領域が重なる部分を返す（辺が接しているだけでも重なりとする）
     * @param {Object} blockA - 領域（x0, y0, x1, y1）
     * @param {Object} blockB - 領域（x0, y0, x1, y1）
     * @returns {Object|null} 重なる部分（無ければ null）
     */
    function getBlockOverlap(blockA, blockB) {
        var overlapX0 = Math.max(blockA.x0, blockB.x0);
        var overlapX1 = Math.min(blockA.x1, blockB.x1);
        var overlapY0 = Math.max(blockA.y0, blockB.y0);
        var overlapY1 = Math.min(blockA.y1, blockB.y1);

        if (overlapX0 <= overlapX1 && overlapY0 <= overlapY1) {
            return { x0: overlapX0, y0: overlapY0, x1: overlapX1, y1: overlapY1 };
        }
        return null;
    }

    /**
     * 領域が別の領域をすべて含むかを返す
     * @param {Object} outerBlock - 外側の領域
     * @param {Object} innerBlock - 内側の領域
     * @returns {boolean} 含むなら true
     */
    function blockContains(outerBlock, innerBlock) {
        return outerBlock.x0 <= innerBlock.x0 && outerBlock.y0 <= innerBlock.y0 && innerBlock.x1 <= outerBlock.x1 && innerBlock.y1 <= outerBlock.y1;
    }

    /**
     * 重なる2つの領域を、重なりの分だけ互いに広げる（どちらかが他方を含むときは何もしない）
     * @param {Object|null} blockA - 領域
     * @param {Object|null} blockB - 領域
     * @returns {void}
     */
    function expandOverlappingBlocks(blockA, blockB) {
        if (!blockA || !blockB || !getBlockOverlap(blockA, blockB)) return;
        if (blockContains(blockA, blockB) || blockContains(blockB, blockA)) return;

        expandBlockInto(blockA, blockB);
        expandBlockInto(blockB, blockA);
    }

    /**
     * 基準の領域に接する辺の分だけ、対象の領域を広げる
     * @param {Object} baseBlock - 基準の領域
     * @param {Object} targetBlock - 広げる領域
     * @returns {void}
     */
    function expandBlockInto(baseBlock, targetBlock) {
        if (baseBlock.x0 <= targetBlock.x0 && targetBlock.x1 <= baseBlock.x1 && targetBlock.y0 <= baseBlock.y1) {
            targetBlock.y0 = Math.min(baseBlock.y0, targetBlock.y0);
            targetBlock.y1 = Math.max(baseBlock.y1, targetBlock.y1);
        }
        if (baseBlock.y0 <= targetBlock.y0 && targetBlock.y1 <= baseBlock.y1 && targetBlock.x0 <= baseBlock.x1) {
            targetBlock.x0 = Math.min(baseBlock.x0, targetBlock.x0);
            targetBlock.x1 = Math.max(baseBlock.x1, targetBlock.x1);
        }
    }

    /**
     * オブジェクトの見た目の範囲を返す。ポイント文字はアウトライン、クリップグループはマスクの内側で測る
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]|undefined} 範囲 [左, 上, 右, 下]（中身の無いグループは undefined）
     */
    function getItemBounds(item) {
        if ("GroupItem" === item.typename) return getGroupBounds(item);

        if ("TextFrame" === item.constructor.name && TextType.AREATEXT !== item.kind) {
            var outlinedText = item.duplicate().createOutline();
            var outlinedBounds = outlinedText.visibleBounds;
            outlinedText.remove();
            return outlinedBounds;
        }

        return item.visibleBounds;
    }

    /**
     * グループの見た目の範囲を返す
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {number[]|undefined} 範囲 [左, 上, 右, 下]
     */
    function getGroupBounds(groupItem) {
        var children = groupItem.pageItems;
        var childBounds = [];
        var clipBounds;

        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            if (child.hasOwnProperty("clipping") && true === child.clipping) {
                clipBounds = child.geometricBounds;
            } else {
                childBounds.push(getItemBounds(child));
            }
        }

        var contentBounds = combineBounds(childBounds);
        var isClippingGroup = groupItem.hasOwnProperty("clipped") && true === groupItem.clipped;
        return isClippingGroup ? intersectBounds(clipBounds, contentBounds) : contentBounds;
    }

    /**
     * すべての範囲を囲む範囲を返す
     * @param {number[][]} boundsList - 範囲 [左, 上, 右, 下] の配列
     * @returns {number[]|undefined} 囲む範囲（空の配列なら undefined）
     */
    function combineBounds(boundsList) {
        var combinedBounds = boundsList[0];
        for (var i = 1; i < boundsList.length; i++) {
            var bounds = boundsList[i];
            combinedBounds = [
                Math.min(combinedBounds[0], bounds[0]),
                Math.max(combinedBounds[1], bounds[1]),
                Math.max(combinedBounds[2], bounds[2]),
                Math.min(combinedBounds[3], bounds[3])
            ];
        }
        return combinedBounds;
    }

    /**
     * 2つの範囲が重なる部分を返す。片方が無ければもう片方をそのまま返す
     * @param {number[]|undefined} boundsA - 範囲 [左, 上, 右, 下]
     * @param {number[]|undefined} boundsB - 範囲 [左, 上, 右, 下]
     * @returns {number[]|undefined} 重なる部分（重ならなければ undefined）
     */
    function intersectBounds(boundsA, boundsB) {
        if (!boundsA || !boundsB) return boundsA || boundsB;

        var isApart = boundsB[0] > boundsA[2] || boundsB[1] < boundsA[3] || boundsB[2] < boundsA[0] || boundsB[3] > boundsA[1];
        if (isApart) return undefined;

        return [
            Math.max(boundsA[0], boundsB[0]),
            Math.min(boundsA[1], boundsB[1]),
            Math.min(boundsA[2], boundsB[2]),
            Math.max(boundsA[3], boundsB[3])
        ];
    }

    /**
     * 長方形を描く（デバッグ用）
     * @param {Document|Layer|GroupItem} container - 描く場所
     * @param {number[]} bounds - 範囲 [左, 上, 右, 下]
     * @param {Object} [properties] - 長方形に設定するプロパティ
     * @returns {PathItem} 描いた長方形
     */
    function drawRectangle(container, bounds, properties) {
        properties = properties || {};

        var rectangle = container.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], -(bounds[3] - bounds[1]));
        rectangle.filled = true;
        rectangle.stroked = false;

        for (var key in properties) {
            if (properties.hasOwnProperty(key)) rectangle[key] = properties[key];
        }
        return rectangle;
    }

    /**
     * 配列をその場でランダムに並べ替える（Fisher-Yates）
     * @param {Array} things - 並べ替える配列
     * @returns {void}
     */
    function shuffle(things) {
        var i = things.length;
        while (i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = things[i];
            things[i] = things[j];
            things[j] = temp;
        }
    }

    /**
     * 外接する長方形が最小になる回転量を返す。
     * 長方形のオブジェクトなら、辺が水平・垂直に戻る角度になる
     * @param {PageItem} item - 対象のオブジェクト（複製で測るので元は動かない）
     * @returns {number} 回転量（度、小数第2位まで）
     */
    function findRotationByMinimalBounds(item) {
        var workingItem = item.duplicate();
        if (undefined !== workingItem.createOutline) workingItem = workingItem.createOutline();

        var rotationStep = ROTATION_START_STEP;
        var rotationAmount = 0;
        var smallestArea = getBoundingArea(workingItem);

        while (Math.abs(rotationStep) >= ROTATION_CONVERGENCE_THRESHOLD) {
            workingItem.rotate(rotationStep);
            var newArea = getBoundingArea(workingItem);

            if (newArea < smallestArea) {
                smallestArea = newArea;
                rotationAmount -= rotationStep;
                rotationStep *= 0.5;
            } else {
                workingItem.rotate(-rotationStep);
                rotationStep *= -0.5;
            }
        }

        workingItem.remove();
        return Math.round(rotationAmount * 100) / 100;
    }

    /**
     * オブジェクトの外接する長方形の面積を返す
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number} 面積
     */
    function getBoundingArea(item) {
        return item.width * item.height;
    }

    /**
     * 間隔・余白の欄の文字列を pt の数値にする（例 "10 mm"）。単位が無ければ pt として読む
     * @param {string|number} lengthValue - 欄の文字列（数値はすでに pt として扱う）
     * @returns {number} pt の値（読めなければ NaN）
     */
    function parseLengthField(lengthValue) {
        if ("number" === typeof lengthValue) return lengthValue;

        var lengthParts = String(lengthValue).match(/^\s*(-?[\d.]+)\s*(\S*)\s*$/);
        if (!lengthParts) return NaN;

        return unitToPt(Number(lengthParts[1]), findUnitByLabel(lengthParts[2]) || UNITS[2]);
    }

    // =========================================
    // メインのダイアログ / Main dialog
    // =========================================

    /**
     * 設定のダイアログボックスを表示する
     * @param {Object} packSettings - 配置の設定（閉じるときに入力値を反映する）
     * @returns {number} [実行] なら 1、[キャンセル] なら 2
     */
    function showPackDialog(packSettings) {
        if (undefined === packSettings.maxAttemptCount) packSettings.maxAttemptCount = getMaxAttemptCount(packSettings.items.length);

        var packDialog = new Window("dialog", getLabel("dialog.title"), undefined, { closeButton: false });
        setupWindow(packDialog);

        var introText = packDialog.add("statictext", undefined, getLabel("message.intro", {
            items: packSettings.items.length,
            artboards: packSettings.doc.artboards.length
        }));
        introText.justify = "center";

        var dialogControls = {};
        addSpacingPanel(packDialog, packSettings, dialogControls);
        addAttemptsPanel(packDialog, packSettings, dialogControls);
        addRotationPanel(packDialog, packSettings, dialogControls);
        dialogControls.showResultsCheckbox = addOptionCheckbox(packDialog, "checkbox.showResults", "tooltip.showResults", packSettings.showResults);

        var buttonRow = addButtonRow(packDialog);
        var btnRandom = buttonRow.leftGroup.add("button", undefined, getLabel("button.random"));
        btnRandom.helpTip = getLabel("tooltip.random");
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnPack = buttonRow.rightGroup.add("button", undefined, getLabel("button.pack"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        var hasRandomResult = false;

        /**
         * ［ランダムに試す］の配置が残っていれば取り消す
         * @returns {void}
         */
        function undoRandomResult() {
            if (!hasRandomResult) return;
            app.undo();
            hasRandomResult = false;
        }

        btnRandom.onClick = function () {
            undoRandomResult();
            applyDialogSettings(dialogControls, packSettings);
            packAndReport(packSettings, true);
            hasRandomResult = true;
            app.redraw();
        };

        btnPack.onClick = function () {
            undoRandomResult();
            applyDialogSettings(dialogControls, packSettings);
            packDialog.close(1);
        };

        addVersionFooter(packDialog);
        prepareDialogWindow(packDialog, SCRIPT_NAME);
        return packDialog.show();
    }

    /**
     * 「間隔」パネルを追加する
     * @param {Window} packDialog - 追加先のダイアログボックス
     * @param {Object} packSettings - 配置の設定
     * @param {Object} dialogControls - 作った入力欄を控えるオブジェクト
     * @returns {void}
     */
    function addSpacingPanel(packDialog, packSettings, dialogControls) {
        var spacingPanel = packDialog.add("panel", undefined, getLabel("panel.spacing"));
        setupPanel(spacingPanel, 6);
        dialogControls.paddingInput = addNumericField(spacingPanel, "padding", String(packSettings.padding), { unit: LENGTH_FIELD_UNIT });
        dialogControls.marginInput = addNumericField(spacingPanel, "margin", String(packSettings.margin), { unit: LENGTH_FIELD_UNIT });
    }

    /**
     * 「試行」パネルを追加する
     * @param {Window} packDialog - 追加先のダイアログボックス
     * @param {Object} packSettings - 配置の設定
     * @param {Object} dialogControls - 作った入力欄を控えるオブジェクト
     * @returns {void}
     */
    function addAttemptsPanel(packDialog, packSettings, dialogControls) {
        var attemptsPanel = packDialog.add("panel", undefined, getLabel("panel.attempts"));
        setupPanel(attemptsPanel, 6);
        dialogControls.maxAttemptsInput = addNumericField(attemptsPanel, "maxAttempts", String(packSettings.maxAttemptCount), { integer: true, min: 1 });
        dialogControls.bestFitDropdown = addBestFitDropdown(attemptsPanel, packSettings.bestFitBy);
        dialogControls.tryHarderCheckbox = addOptionCheckbox(attemptsPanel, "checkbox.tryHarder", "tooltip.tryHarder", packSettings.tryHarder);
        dialogControls.doNotSortCheckbox = addOptionCheckbox(attemptsPanel, "checkbox.doNotSort", "tooltip.doNotSort", packSettings.doNotSort);
    }

    /**
     * 「回転」パネルを追加する
     * @param {Window} packDialog - 追加先のダイアログボックス
     * @param {Object} packSettings - 配置の設定
     * @param {Object} dialogControls - 作った入力欄を控えるオブジェクト
     * @returns {void}
     */
    function addRotationPanel(packDialog, packSettings, dialogControls) {
        var rotationPanel = packDialog.add("panel", undefined, getLabel("panel.rotation"));
        setupPanel(rotationPanel, 6);
        dialogControls.allow90RotationCheckbox = addOptionCheckbox(rotationPanel, "checkbox.allow90Rotation", "tooltip.allow90Rotation", packSettings.allow90DegreeRotation);
        dialogControls.allowAnyRotationCheckbox = addOptionCheckbox(rotationPanel, "checkbox.allowAnyRotation", "tooltip.allowAnyRotation", packSettings.allowAnyRotation);
    }

    /**
     * ステップボタン付きの数値欄を追加する（項目名と tooltip は同じキーで LABELS から引く）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel と LABELS.tooltip のキー
     * @param {string} fieldText - 欄の初期値
     * @param {Object} stepOptions - unit / integer / min（addSteppedField と同じ）
     * @returns {EditText} 入力欄
     */
    function addNumericField(parentPanel, labelKey, fieldText, stepOptions) {
        stepOptions.label = labelText("fieldLabel." + labelKey);
        stepOptions.labelWidth = FIELD_LABEL_WIDTH;
        stepOptions.text = fieldText;
        stepOptions.characters = FIELD_CHARACTERS;
        stepOptions.step = 1;

        var numberInput = addSteppedField(parentPanel, stepOptions);
        numberInput.helpTip = getLabel("tooltip." + labelKey);
        numberInput.fieldLabel.helpTip = numberInput.helpTip;
        return numberInput;
    }

    /**
     * 「優先」のドロップダウンリストを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} bestFitBy - 初期値（"count" または "area"）
     * @returns {DropDownList} ドロップダウンリスト
     */
    function addBestFitDropdown(parentPanel, bestFitBy) {
        var bestFitRow = parentPanel.add("group");
        setupRow(bestFitRow, "left", STEPPER_FIELD_SPACING + STEPPER_SIDE_MARGIN); /* 数値欄の左端にそろえる / line up with the numeric fields */

        var bestFitLabel = bestFitRow.add("statictext", undefined, labelText("fieldLabel.bestFit"));
        bestFitLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        bestFitLabel.justify = "right";
        bestFitLabel.helpTip = getLabel("tooltip.bestFit");

        var bestFitDropdown = bestFitRow.add("dropdownlist", undefined, [getLabel("dropdown.count"), getLabel("dropdown.area")]);
        bestFitDropdown.selection = ("area" === bestFitBy) ? 1 : 0;
        bestFitDropdown.helpTip = bestFitLabel.helpTip;
        return bestFitDropdown;
    }

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Group|Panel|Window} parentContainer - 追加先
     * @param {string} labelPath - 文言の LABELS のパス
     * @param {string} tooltipPath - tooltip の LABELS のパス
     * @param {boolean} isChecked - 初期状態
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parentContainer, labelPath, tooltipPath, isChecked) {
        var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.value = isChecked;
        optionCheckbox.helpTip = getLabel(tooltipPath);
        return optionCheckbox;
    }

    /**
     * ダイアログボックスの入力値を設定に写す
     * @param {Object} dialogControls - ダイアログボックスの入力欄
     * @param {Object} packSettings - 書き込み先の設定
     * @returns {void}
     */
    function applyDialogSettings(dialogControls, packSettings) {
        /* ボタンを押しても欄の確定が走らないことがあるので、計算式や単位の換算を先に済ませる
           A button press may not commit the field, so evaluate expressions and units first */
        dialogControls.paddingInput.onChange();
        dialogControls.marginInput.onChange();
        dialogControls.maxAttemptsInput.onChange();

        packSettings.padding = dialogControls.paddingInput.text;
        packSettings.margin = dialogControls.marginInput.text;
        packSettings.maxAttemptCount = parseInt(dialogControls.maxAttemptsInput.text, 10);
        packSettings.bestFitBy = (1 === dialogControls.bestFitDropdown.selection.index) ? "area" : "count";
        packSettings.allow90DegreeRotation = dialogControls.allow90RotationCheckbox.value;
        packSettings.allowAnyRotation = dialogControls.allowAnyRotationCheckbox.value;
        packSettings.tryHarder = dialogControls.tryHarderCheckbox.value;
        packSettings.doNotSort = dialogControls.doNotSortCheckbox.value;
        packSettings.showResults = dialogControls.showResultsCheckbox.value;
    }

    // =========================================
    // 進捗と結果 / Progress and results
    // =========================================

    /**
     * 進捗ウィンドウを作る
     * @returns {Window} setItemsPackedProgress / setAttemptProgress / setBestAttempt を持つウィンドウ
     */
    function makeProgressWindow() {
        var progressWindow = new Window("window", getLabel("dialog.progress"), undefined, { closeButton: false });
        setupWindow(progressWindow);

        var packedRow = addProgressRow(progressWindow, getLabel("progress.packed"));
        var attemptRow = addProgressRow(progressWindow, getLabel("progress.attempt"));

        progressWindow.setItemsPackedProgress = function (packedItemCount, totalItemCount) {
            updateProgressRow(packedRow, packedItemCount, totalItemCount);
        };

        progressWindow.setAttemptProgress = function (attemptNumber, maxAttemptCount) {
            updateProgressRow(attemptRow, attemptNumber, maxAttemptCount);
        };

        progressWindow.setBestAttempt = function (binCount, attemptIndex) {
            packedRow.caption.text = getLabel("progress.bestAttempt", { attempt: attemptIndex + 1, artboards: binCount });
        };

        return progressWindow;
    }

    /**
     * 見出し・進捗バー・件数のひと組を追加する
     * @param {Window} progressWindow - 追加先のウィンドウ
     * @param {string} captionText - 見出しの文言
     * @returns {{caption: StaticText, bar: Progressbar, count: StaticText}} 追加した部品
     */
    function addProgressRow(progressWindow, captionText) {
        var progressGroup = progressWindow.add("group");
        progressGroup.orientation = "column";
        progressGroup.alignChildren = "fill";

        var caption = progressGroup.add("statictext", undefined, captionText);

        var barRow = progressGroup.add("group");
        setupRow(barRow, "fill");
        var progressBar = barRow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;
        var countText = barRow.add("statictext", undefined, "0 / 0");
        countText.preferredSize.width = PROGRESS_COUNT_WIDTH;

        return { caption: caption, bar: progressBar, count: countText };
    }

    /**
     * 進捗バーと件数を更新して描き直す
     * @param {{caption: StaticText, bar: Progressbar, count: StaticText}} progressRow - addProgressRow() の戻り値
     * @param {number} currentValue - 今の値
     * @param {number} maxValue - 最大値
     * @returns {void}
     */
    function updateProgressRow(progressRow, currentValue, maxValue) {
        progressRow.bar.maxvalue = maxValue;
        progressRow.bar.value = currentValue;
        progressRow.count.text = currentValue + " / " + maxValue;
        progressRow.bar.window.update();
    }

    /**
     * 配置の結果のダイアログボックスを表示する
     * @param {Attempt} attempt - 採用した試行
     * @returns {void}
     */
    function showResultsDialog(attempt) {
        var resultsDialog = new Window("dialog", getLabel("dialog.results"), undefined, { closeButton: false });
        setupWindow(resultsDialog);

        var resultText = resultsDialog.add("statictext", undefined, buildResultLines(attempt).join("\n"), { multiline: true });
        resultText.preferredSize = RESULT_TEXT_SIZE;

        var buttonRow = addButtonRow(resultsDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.close"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        app.redraw();
        prepareDialogWindow(resultsDialog, SCRIPT_NAME + "_results");
        resultsDialog.show();
    }

    /**
     * 結果として表示する行を返す
     * @param {Attempt} attempt - 採用した試行
     * @returns {string[]} 表示する行
     */
    function buildResultLines(attempt) {
        var remainingCount = attempt.remainingBlocks.length;
        var summaryLine = (0 === remainingCount)
            ? getLabel("result.success", { count: attempt.packedBlocks.length })
            : getLabel("result.failure", { count: remainingCount });

        return [
            summaryLine,
            "",
            labelValueText("result.attemptUsed", attempt.index + 1),
            labelValueText("result.sortMethod", getLabel(LABELS.sortMethod[attempt.sortType || "none"])),
            labelValueText("result.score", Math.round(attempt.score)),
            ""
        ].concat(attempt.info);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * オブジェクトが選択されているかを返す（文字の選択は対象外）
     * @returns {boolean} 配置できる選択があれば true
     */
    function hasPackableSelection() {
        if (0 === app.documents.length) return false;
        var selectedItems = app.activeDocument.selection;
        return selectedItems.length > 0 && "TextRange" !== selectedItems.typename;
    }

    /**
     * ダイアログボックスで設定を受け取り、配置を実行する
     * @returns {void}
     */
    function main() {
        if (!hasPackableSelection()) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var packSettings = DEFAULT_PACK_SETTINGS;
        packSettings.doc = app.activeDocument;
        packSettings.items = packSettings.doc.selection;

        if (packSettings.showUI && 2 === showPackDialog(packSettings)) return;

        var progressWindow = makeProgressWindow();
        packSettings.progressWindow = progressWindow;
        progressWindow.center();
        progressWindow.show();

        packAndReport(packSettings, false);
    }

    main();

})();
