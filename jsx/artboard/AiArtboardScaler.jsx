#target illustrator
#targetengine "AiArtboardScalerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードを現在のサイズを基準にスケール変更します。ダイアログを開いたままライブプレビューできます。
対象は現在のアートボード／すべて／指定から選べ、9つの基準点のいずれかを固定して再計算します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiArtboardScaler.md

### Overview

Scales artboards relative to their current size, with a live preview while the dialog stays open.
The target can be the current artboard, all of them, or a specified set, recalculated around one of nine reference points.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiArtboardScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiArtboardScaler";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiArtboardScaler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiArtboardScaler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

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

    var OPTION_PANEL_SPACING = 6;            /* 対象・サイズパネル内の要素間隔 / spacing inside the target and size panels */
    var FIELD_LABEL_WIDTH = 64;              /* ラベル幅を揃えるための固定幅（「スケール:」が収まる幅）/ fixed label width (fits "Scale:") */
    var NUMBER_FIELD_CHARS = 4;              /* スケール・幅・高さ欄の文字数 / width of the scale, width and height fields */
    var SELECTION_FIELD_CHARS = 8;           /* 「指定」欄の文字数（通常の2倍）/ width of the Specify field (twice the usual) */
    var SCALE_COLUMN_GAP = 16;               /* 入力欄の列と基準点の列の間隔 / gap between the field column and the anchor column */
    var CHECKBOX_GROUP_MARGINS = [0, 10, 0, 0]; /* チェックボックス群の余白（上に10px）/ checkbox group margins (10px on top) */

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 数値を小数2桁に丸めて文字列で返す / Round a number to 2 decimals and return as string
     * @param {number} value - 丸める値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return "" + (Math.round(value * 100) / 100);
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
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
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
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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
            title: { ja: "アートボードサイズ変更", en: "Resize Artboards" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            scale:  { ja: "サイズ・スケール", en: "Size & Scale" }
        },
        radio: {
            current: { ja: "現在のアートボード", en: "Current artboard" },
            all:     { ja: "すべてのアートボード", en: "All artboards" },
            specify: { ja: "指定", en: "Specify" }
        },
        checkbox: {
            scaleObjects: { ja: "オブジェクトと一緒に拡大・縮小", en: "Scale objects together" },
            pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize for pixel grid" }
        },
        fieldLabel: {
            scale:  { ja: "スケール", en: "Scale" },
            width:  { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            anchor: { ja: "基準点", en: "Anchor" }
        },
        tooltip: {
            current: { ja: "いま選ばれているアートボードだけを変更します。", en: "Changes only the artboard that is currently active." },
            all:     { ja: "ドキュメント内のすべてのアートボードを変更します。", en: "Changes every artboard in the document." },
            specify: { ja: "番号で対象を指定します（例: 3, 4 または 3-5）。", en: "Picks the artboards by number (for example 3, 4 or 3-5)." },
            scale:   { ja: "現在のサイズに対する倍率（％）です。幅・高さと連動します。", en: "Percentage of the current size. It is linked to the width and height." },
            size:    { ja: "変更後のサイズです。入力するとスケールが連動して変わります。", en: "The size after resizing. Typing here updates the scale." },
            anchor:  { ja: "サイズを変えるときに動かさない位置です。3×3のマスで選びます。", en: "The point that stays put while the artboard is resized. Pick it on the 3x3 grid." },
            scaleObjects: {
                ja: "アートボードの拡大・縮小に合わせて、載っているオブジェクトも一緒に変形します。",
                en: "Scales the objects on the artboard along with the artboard itself."
            },
            pixelGrid: {
                ja: "アートボードの位置とサイズを整数ピクセルにそろえます。",
                en: "Snaps the artboard position and size to whole pixels."
            },
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
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply:  { ja: "適用", en: "Apply" }
        },
        alert: {
            noDocument:      { ja: "開いているドキュメントがありません。", en: "No documents are open." },
            invalidNumber:   { ja: "正の数値を入力してください。", en: "Please enter positive numbers." },
            invalidSelection: {
                ja: "対象アートボードの指定が正しくありません（例: 3, 4 または 3-5）。",
                en: "Invalid artboard selection (e.g. 3, 4 or 3-5)."
            },
            transformError: {
                ja: "変形の適用に失敗したため、処理を中止して元に戻しました。",
                en: "Failed to apply the transform; the operation was cancelled and reverted."
            },
            restoreError: {
                ja: "アートボードの復元に失敗しました。手動で元に戻してください（取り消し等）。",
                en: "Failed to restore the artboards. Please revert manually (e.g. Undo)."
            }
        }
    };

    // =========================================
    // 対象指定の解析 / Target spec parsing
    // =========================================

    /**
     * 前後の空白を除去する（ES3にString.trimが無いため） / Trim whitespace (ES3 has no String.trim)
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(sourceText) {
        return ("" + sourceText).replace(/^\s+/, "").replace(/\s+$/, "");
    }

    /**
     * 「3, 4」「3-5」形式の指定を0始まりのアートボード索引配列に厳密変換する
     * Strictly parse a "3, 4" / "3-5" style spec into 0-based artboard indices
     * 各トークンを正規表現で完全一致検証し、"3abc"・"1-2-3"・空トークン(",")等は不正扱い。
     * @param {string} selectionText - 入力文字列（1始まり）
     * @param {number} artboardCount - アートボード総数
     * @returns {number[]|null} 索引配列。不正な場合は null
     */
    function parseArtboardSelection(selectionText, artboardCount) {
        var singlePattern = /^\d+$/;                 /* 単一番号 / single number */
        var rangePattern = /^\d+\s*-\s*\d+$/;        /* 範囲（前後の空白可） / range (spaces allowed) */
        var artboardIndices = [];
        var seenIndices = {};
        var tokens = ("" + selectionText).split(",");
        for (var i = 0; i < tokens.length; i++) {
            var token = trimWhitespace(tokens[i]);
            /* 空トークン（"1," ",2" "1,,2" 等）は不正扱い / empty token (from "1," ",2" "1,,2") is invalid */
            if (token === "") { return null; }

            var rangeStart, rangeEnd;
            if (rangePattern.test(token)) {
                var dashPosition = token.indexOf("-");
                rangeStart = parseInt(trimWhitespace(token.substring(0, dashPosition)), 10);
                rangeEnd = parseInt(trimWhitespace(token.substring(dashPosition + 1)), 10);
            } else if (singlePattern.test(token)) {
                rangeStart = rangeEnd = parseInt(token, 10);
            } else {
                return null; /* 形式不一致（"3abc" "1-2-3" 等） / format mismatch */
            }
            if (rangeStart > rangeEnd) { var swappedStart = rangeStart; rangeStart = rangeEnd; rangeEnd = swappedStart; } /* 逆順は入れ替え / swap reversed range */

            for (var artboardNumber = rangeStart; artboardNumber <= rangeEnd; artboardNumber++) {
                if (artboardNumber < 1 || artboardNumber > artboardCount) { return null; } /* 範囲外は不正 / out of range is invalid */
                var artboardIndex = artboardNumber - 1;
                if (!seenIndices[artboardIndex]) { seenIndices[artboardIndex] = true; artboardIndices.push(artboardIndex); } /* 重複は1回だけ / dedupe */
            }
        }
        return artboardIndices.length ? artboardIndices : null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 「ラベル: [入力欄] 単位」の1行を生成する / Build one "label: [input] unit" row
     * @param {Group} parentGroup - 追加先
     * @param {string} captionText - コロン付きの項目名
     * @param {string} defaultValue - 入力欄の初期値
     * @param {string} unitLabel - 入力欄の後ろに表示する単位
     * @param {string} [tooltipText] - 入力欄に付けるツールチップ
     * @returns {EditText} 生成した入力欄
     */
    function addSizeField(parentGroup, captionText, defaultValue, unitLabel, tooltipText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        var captionLabel = fieldRow.add("statictext", undefined, captionText);
        captionLabel.preferredSize.width = FIELD_LABEL_WIDTH; /* ラベル幅を固定して揃える / Fix width to align labels */
        captionLabel.justify = "right";                       /* 右揃え / Right-align the label */
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var valueInput;
        /* 増減後は onChanging を呼んで連動欄とプレビューを更新する / fire onChanging so linked fields and the preview follow */
        var stepperGroup = addStepper(stepperInputGroup, function () { return valueInput; }, {
            min: 0,
            onStep: function (numberInput) {
                if (typeof numberInput.onChanging === "function") numberInput.onChanging();
            }
        });
        valueInput = stepperInputGroup.add("edittext", undefined, defaultValue);
        if (tooltipText) valueInput.helpTip = tooltipText;
        valueInput.characters = NUMBER_FIELD_CHARS;
        bindSteppedArrowKeys(valueInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
        fieldRow.add("statictext", undefined, unitLabel); /* 入力欄の後ろに単位 / Unit after the input */
        return valueInput;
    }

    /**
     * プレビュー更新を要求する（コントローラ未接続なら何もしない） / Request a preview refresh (no-op until wired)
     * @param {Window} resizeDialog - onPreview を持つダイアログ
     * @returns {void}
     */
    function requestPreview(resizeDialog) {
        if (resizeDialog.onPreview) { resizeDialog.onPreview(); }
    }

    /**
     * 対象パネル（現在のアートボード／すべてのアートボード／指定）を構築する
     * Build the target panel (Current artboard / All artboards / Specify)
     * @param {Window} resizeDialog - 追加先ダイアログ（コントロールをプロパティとして公開する）
     * @param {string} defaultSelection - 指定入力欄の初期値（例: "1-6"）
     * @returns {void}
     */
    function addTargetPanel(resizeDialog, defaultSelection) {
        var targetPanel = resizeDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, OPTION_PANEL_SPACING);

        /**
         * パネルに左寄せの行を追加する
         * @returns {Group} 追加した行
         */
        function addLeftRow() {
            var radioRow = targetPanel.add("group");
            radioRow.orientation = "row";
            radioRow.alignment = "left";
            return radioRow;
        }

        /* 「現在のアートボード」ラジオ / "Current artboard" radio */
        var currentRadio = addLeftRow().add("radiobutton", undefined, getLabel("radio.current"));
        currentRadio.helpTip = getLabel("tooltip.current");

        /* 「すべてのアートボード」ラジオ / "All artboards" radio */
        var allRadio = addLeftRow().add("radiobutton", undefined, getLabel("radio.all"));
        allRadio.helpTip = getLabel("tooltip.all");

        /* 「指定」ラジオ＋範囲入力 / "Specify" radio with range input */
        var specifyRow = addLeftRow();
        var specifyRadio = specifyRow.add("radiobutton", undefined, getLabel("radio.specify"));
        specifyRadio.helpTip = getLabel("tooltip.specify");
        var selectInput = specifyRow.add("edittext", undefined, defaultSelection);
        selectInput.helpTip = getLabel("tooltip.specify");
        selectInput.characters = SELECTION_FIELD_CHARS;

        var targetRadios = [currentRadio, allRadio, specifyRadio];

        /**
         * ラジオは親が異なると排他にならないため手動で同期する / Sync manually since radios in different parents are not exclusive
         * @param {RadioButton} activeRadio - 選ばれたラジオ
         * @returns {void}
         */
        function selectTarget(activeRadio) {
            for (var i = 0; i < targetRadios.length; i++) { targetRadios[i].value = (targetRadios[i] === activeRadio); }
            selectInput.enabled = (activeRadio === specifyRadio); /* 入力欄は「指定」時のみ有効 / input enabled only for "Specify" */
            requestPreview(resizeDialog);
        }
        currentRadio.onClick = function () { selectTarget(currentRadio); };
        allRadio.onClick = function () { selectTarget(allRadio); };
        specifyRadio.onClick = function () { selectTarget(specifyRadio); };
        selectInput.onChanging = function () { requestPreview(resizeDialog); };

        /* 初期状態は「すべてのアートボード」 / Default to "All artboards" */
        allRadio.value = true;
        selectInput.enabled = false;

        resizeDialog.currentRadio = currentRadio;
        resizeDialog.allRadio = allRadio;
        resizeDialog.specifyRadio = specifyRadio;
        resizeDialog.selectInput = selectInput;
    }

    /**
     * サイズ・スケール統合パネルを構築する / Build the merged size & scale panel
     * スケール%を一意の倍率とし、幅・高さ（現在の定規単位）はアクティブアートボードの現在サイズ基準で相互連動する。
     * The scale % is a single uniform ratio; width/height (in the current ruler unit) are linked to the active artboard's current size.
     * @param {Window} resizeDialog - 追加先ダイアログ（コントロールをプロパティとして公開する）
     * @param {string} unitLabel - 幅・高さの単位ラベル（例: "mm"）
     * @param {number} baseWidth - 幅の基準値：アクティブアートボードの現在幅を現在の定規単位へ変換済み（ptではない）
     * @param {number} baseHeight - 高さの基準値：アクティブアートボードの現在高さを現在の定規単位へ変換済み（ptではない）
     * @returns {void}
     */
    function addScalePanel(resizeDialog, unitLabel, baseWidth, baseHeight) {
        var scalePanel = resizeDialog.add("panel", undefined, getLabel("panel.scale"));
        setupPanel(scalePanel, OPTION_PANEL_SPACING);

        /* 2列レイアウト：左=スケール/幅/高さの3行、右=基準点グリッド / Two columns: left has scale/width/height rows, right holds the anchor grid */
        var columnsRow = scalePanel.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["left", "center"];
        columnsRow.spacing = SCALE_COLUMN_GAP;

        var fieldColumn = columnsRow.add("group");
        fieldColumn.orientation = "column";
        fieldColumn.alignChildren = ["left", "top"];
        fieldColumn.spacing = OPTION_PANEL_SPACING;
        resizeDialog.scaleInput = addSizeField(fieldColumn, labelText("fieldLabel.scale"), "100", "%", getLabel("tooltip.scale"));
        resizeDialog.widthInput = addSizeField(fieldColumn, labelText("fieldLabel.width"), formatNumber(baseWidth), unitLabel, getLabel("tooltip.size"));
        resizeDialog.heightInput = addSizeField(fieldColumn, labelText("fieldLabel.height"), formatNumber(baseHeight), unitLabel, getLabel("tooltip.size"));

        /* 基準点グリッド（2列目） / Anchor reference-point grid (second column) */
        addAnchorColumn(resizeDialog, columnsRow);

        /* チェックボックス群（上に10pxの余白） / Checkbox group (10px top margin) */
        var checkboxGroup = scalePanel.add("group");
        checkboxGroup.orientation = "column";
        checkboxGroup.alignChildren = ["left", "top"];
        checkboxGroup.margins = CHECKBOX_GROUP_MARGINS;

        /* オブジェクトも一緒に拡大・縮小するか / Whether to scale the objects along with the artboard */
        var scaleObjectsCheckbox = checkboxGroup.add("checkbox", undefined, getLabel("checkbox.scaleObjects"));
        scaleObjectsCheckbox.helpTip = getLabel("tooltip.scaleObjects");
        scaleObjectsCheckbox.value = true;

        /* アートボードのX/Y/W/Hを整数化してピクセルグリッドに合わせる / Round artboard X/Y/W/H to integers for the pixel grid */
        var pixelGridCheckbox = checkboxGroup.add("checkbox", undefined, getLabel("checkbox.pixelGrid"));
        pixelGridCheckbox.helpTip = getLabel("tooltip.pixelGrid");
        pixelGridCheckbox.value = false;

        linkScaleFields(resizeDialog, baseWidth, baseHeight);
        scaleObjectsCheckbox.onClick = function () {
            requestPreview(resizeDialog);
        };
        pixelGridCheckbox.onClick = function () {
            requestPreview(resizeDialog);
        };

        resizeDialog.scaleObjectsCheckbox = scaleObjectsCheckbox;
        resizeDialog.pixelGridCheckbox = pixelGridCheckbox;
    }

    /**
     * スケール・幅・高さの3欄を連動させる（入力中は相互に更新し、確定時に正規化する）
     * Link the scale, width and height fields (update each other while typing, normalize on commit)
     * @param {Window} resizeDialog - scaleInput / widthInput / heightInput を持つダイアログ
     * @param {number} baseWidth - 幅の基準値（定規単位）
     * @param {number} baseHeight - 高さの基準値（定規単位）
     * @returns {void}
     */
    function linkScaleFields(resizeDialog, baseWidth, baseHeight) {
        var scaleInput = resizeDialog.scaleInput;
        var widthInput = resizeDialog.widthInput;
        var heightInput = resizeDialog.heightInput;

        /**
         * スケール%から幅・高さ(現在サイズ×%)を再計算する / Recalc width/height (current size × %) from the scale
         * @returns {void}
         */
        function applyScaleToSize() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent)) { return; }
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
        }

        /**
         * 編集中の寸法欄からスケール%を逆算し、スケール欄ともう一方の寸法欄だけ更新する
         * Derive the scale from the edited size field, updating only the scale field and the OTHER size field
         * （編集中の欄自身は書き換えない＝小数点入力が消える不具合を防ぐ / never rewrite the field being edited, so decimals can be typed）
         * @param {number} editedValue - 編集中の欄の値
         * @param {number} editedBase - 編集中の欄の基準サイズ
         * @param {EditText} otherField - もう一方（連動更新する）寸法欄
         * @param {number} otherBase - もう一方の基準サイズ
         * @returns {void}
         */
        function applySizeToScale(editedValue, editedBase, otherField, otherBase) {
            if (isNaN(editedValue) || editedBase === 0) { return; }
            var percent = editedValue / editedBase * 100;
            scaleInput.text = formatNumber(percent);
            otherField.text = formatNumber(otherBase * percent / 100);
        }

        /* 直近の有効なスケール%（入力強化の復帰先） / Last valid scale % (fallback for input hardening) */
        var lastValidScale = 100;

        /**
         * 確定時に3欄をスケール基準へ正規化する（不正値は直近の有効値へ復帰） / On commit, canonicalize all three fields to the scale (invalid → last valid)
         * @returns {void}
         */
        function normalizeFields() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent) || percent <= 0) {
                percent = lastValidScale; /* 空・0・負・非数値は直近の有効値へ / empty/0/negative/NaN falls back */
            } else {
                lastValidScale = percent;
            }
            scaleInput.text = formatNumber(percent);
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
            requestPreview(resizeDialog);
        }

        scaleInput.onChanging = function () {
            applyScaleToSize();
            requestPreview(resizeDialog);
        };
        widthInput.onChanging = function () {
            applySizeToScale(parseFloat(widthInput.text), baseWidth, heightInput, baseHeight);
            requestPreview(resizeDialog);
        };
        heightInput.onChanging = function () {
            applySizeToScale(parseFloat(heightInput.text), baseHeight, widthInput, baseWidth);
            requestPreview(resizeDialog);
        };
        /* 確定（Enter/フォーカスアウト）で正規化 / Normalize on commit (Enter / focus-out) */
        scaleInput.onChange = normalizeFields;
        widthInput.onChange = normalizeFields;
        heightInput.onChange = normalizeFields;
    }

    /**
     * サイズ入力ダイアログを構築する / Build the size-input dialog
     * @param {string} unitLabel - 表示する単位ラベル（例: "mm"）
     * @param {string} defaultWidth - 幅入力欄の初期値
     * @param {string} defaultHeight - 高さ入力欄の初期値
     * @param {number} artboardCount - アートボード総数（選択の初期値に使用）
     * @returns {Window} ダイアログウィンドウ（各入力コントロールを公開）
     */
    function createResizeDialog(unitLabel, defaultWidth, defaultHeight, artboardCount) {
        /* タイトルバーにバージョンを表示 / Show version in the title bar */
        var resizeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(resizeDialog);

        /* 対象パネル / Target panel */
        addTargetPanel(resizeDialog, "1-" + artboardCount);

        /* サイズ・スケール統合パネル（アクティブアートボードの現在サイズを基準に） / Merged size & scale panel (based on the active artboard's current size) */
        addScalePanel(resizeDialog, unitLabel, parseFloat(defaultWidth), parseFloat(defaultHeight));

        /* ボタン類はパネル幅いっぱいには広げず右寄せ / Keep buttons right-aligned, not full width */
        var buttonRow = addButtonRow(resizeDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), {name: "cancel"});
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"), {name: "ok"});
        alignRightOnlyButtonRow(buttonRow);

        return resizeDialog;
    }

    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 1]     : [0.42, 0.42, 0.42, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 1]     : [0.27, 0.27, 0.27, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 0.4]   : [0.42, 0.42, 0.42, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 0.3]   : [0.27, 0.27, 0.27, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

    /**
     * 基準点の列（見出し＋3×3の基準点ウィジェット）を構築する / Build the anchor column (caption + 3×3 anchor widget)
     * 選択に応じて resizeDialog.anchorX / resizeDialog.anchorY に 0/0.5/1 の割合を設定する（既定=左上）。
     * @param {Window} resizeDialog - 割合を書き込むダイアログ
     * @param {Group} parentGroup - 追加先グループ
     * @returns {void}
     */
    function addAnchorColumn(resizeDialog, parentGroup) {
        var anchorColumn = parentGroup.add("group");
        anchorColumn.orientation = "column";
        anchorColumn.alignChildren = ["center", "top"];
        anchorColumn.spacing = 4;
        anchorColumn.add("statictext", undefined, getLabel("fieldLabel.anchor"));

        /* 0..8 行優先（0=左上, 4=中央, 8=右下）。既定=左上 / 0..8 row-major (0=top-left, 4=center, 8=bottom-right); default: top-left */
        var anchorWidget = addAnchorWidget(anchorColumn, 0, function () {
            commitAnchor();
            requestPreview(resizeDialog);
        });
        anchorWidget.helpTip = getLabel("tooltip.anchor");

        /**
         * 選択中のセルを割合(0/0.5/1)に変換して resizeDialog に反映する / Map the selected cell to fractions and store on the dialog
         * @returns {void}
         */
        function commitAnchor() {
            var anchorRatio = getAnchorRatio(getAnchorWidgetIndex(anchorWidget));
            resizeDialog.anchorX = anchorRatio[0];
            resizeDialog.anchorY = anchorRatio[1];
        }
        commitAnchor(); /* 既定=左上 / default: top-left */
    }

    // =========================================
    // オブジェクトの拡縮 / Object scaling
    // =========================================

    /**
     * 1アイテムを逆拡縮して元の位置へ戻す（rollback用） / Inverse-scale one item and put it back (for rollback)
     * @param {PageItem} pageItem - 戻すアイテム
     * @param {number} ratioX - 掛けた横の倍率
     * @param {number} ratioY - 掛けた縦の倍率
     * @param {number[]} originalPosition - 元の position [左, 上]
     * @returns {void}
     */
    function inverseResizeItem(pageItem, ratioX, ratioY, originalPosition) {
        try {
            var inverseX = (1 / ratioX) * 100;
            var inverseY = (1 / ratioY) * 100;
            /* 第7引数(changeLineWidths)は線幅拡縮のパーセント値を渡す（参考実装準拠） / 7th arg is the line-width scale percentage (per the reference) */
            pageItem.resize(inverseX, inverseY, true, true, true, true, inverseX, Transformation.TOPLEFT);
            pageItem.position = [originalPosition[0], originalPosition[1]];
        } catch (e) {}
    }

    /**
     * 【確定用】指定比率でオブジェクト群を基準点(anchor)基準に一発で拡縮する（線幅も拡縮／途中失敗は自前で完全復元）
     * [For commit] Scale items about the anchor once (also scales line width); on mid-way failure, fully reverts its own items
     * artboardsResizeWithObjects.jsx(Alexander Ladygin) を参考にした resize()+position 方式。
     * resize() でサイズ・線幅を拡縮（アイテム左上基準）し、position で基準点からのオフセットを拡縮して再配置する。
     * オブジェクトごとに { pageItem, originalPosition, resized } を記録し、失敗時はこの呼び出しで変形済みの分を確実に戻す。
     * @param {PageItem[]} targetItems - 拡縮するアイテム
     * @param {number} ratioX - 横の倍率
     * @param {number} ratioY - 縦の倍率
     * @param {number} anchorX - 基準点のX
     * @param {number} anchorY - 基準点のY
     * @returns {object} { success:boolean, transformedItems:PageItem[][, error:Error] }
     */
    function resizeItemsAbout(targetItems, ratioX, ratioY, anchorX, anchorY) {
        var transformedItems = [];
        if (!targetItems || !targetItems.length) { return { success: true, transformedItems: transformedItems }; }
        var resizeStates = []; /* {pageItem, originalPosition, resized} 途中失敗時の復元用 / for rollback on failure */
        for (var i = 0; i < targetItems.length; i++) {
            var pageItem = targetItems[i];
            var originalPosition = pageItem.position; // [左, 上]（ドキュメント座標） / [left, top] in document coordinates
            var resizeState = { pageItem: pageItem, originalPosition: [originalPosition[0], originalPosition[1]], resized: false };
            resizeStates.push(resizeState);
            try {
                /* サイズ・線幅を拡縮（アイテム左上基準）。第7引数=線幅拡縮のパーセント値（参考実装準拠） / Scale size & line width about the item's top-left; 7th arg = line-width scale percentage (per the reference) */
                pageItem.resize(ratioX * 100, ratioY * 100, true, true, true, true, ratioX * 100, Transformation.TOPLEFT);
                resizeState.resized = true; /* resize成功。以降のpositionで失敗しても逆resizeで戻せる / resize done; still recoverable if position fails */
                /* 基準点からのオフセットを拡縮して再配置 / Reposition by scaling the offset from the anchor point */
                pageItem.position = [
                    anchorX + (originalPosition[0] - anchorX) * ratioX,
                    anchorY + (originalPosition[1] - anchorY) * ratioY
                ];
                transformedItems.push(pageItem);
            } catch (e) {
                /* 途中失敗：この呼び出しで resize 済み（position前後どちらも）を確実に復元 / mid-way failure: revert every item resized in this call */
                for (var j = resizeStates.length - 1; j >= 0; j--) {
                    if (resizeStates[j].resized) {
                        inverseResizeItem(resizeStates[j].pageItem, ratioX, ratioY, resizeStates[j].originalPosition);
                    }
                }
                return { success: false, transformedItems: [], error: e };
            }
        }
        return { success: true, transformedItems: transformedItems };
    }

    // =========================================
    // オブジェクトの所属 / Object ownership
    // =========================================

    /**
     * 指定アートボード上の編集可能オブジェクトを安全に取得する（selection/active が途中例外でも壊れない）
     * Safely collect editable objects on the given artboard (selection/active stay consistent even if an exception occurs)
     * @param {Document} doc - 対象ドキュメント
     * @param {number} artboardIndex - アートボードの索引
     * @returns {PageItem[]} アートボード上のオブジェクト
     */
    function getObjectsOnArtboard(doc, artboardIndex) {
        var artboardItems = [];
        try {
            doc.selection = null;
            doc.artboards.setActiveArtboardIndex(artboardIndex);
            doc.selectObjectsOnActiveArtboard();
            var currentSelection = doc.selection;
            /* selection が null/未定義相当や配列でない場合も安全に扱う / Handle null / non-array selection safely */
            if (currentSelection && typeof currentSelection.length === "number") {
                for (var i = 0; i < currentSelection.length; i++) { artboardItems.push(currentSelection[i]); }
            }
        } finally {
            /* 例外の有無に関わらず選択を解除する / clear the selection whether or not an exception occurred */
            try { doc.selection = null; } catch (e) {}
        }
        return artboardItems;
    }

    /**
     * オブジェクトの一意識別子を返す（uuid優先、無ければ null） / Return an object's unique id (prefer uuid; null if unavailable)
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {string|null} uuid
     */
    function getItemUuid(pageItem) {
        try {
            if (pageItem.uuid) { return pageItem.uuid; }
        } catch (e) {}
        return null;
    }

    /**
     * 点(x,y)がアートボード矩形[左,上,右,下]の内側か / Whether point (x,y) is inside the artboard rect [L,T,R,B]
     * @param {number[]} artboardRect - アートボードの矩形
     * @param {number} x - 点のX
     * @param {number} y - 点のY
     * @returns {boolean} 内側なら true
     */
    function rectContainsPoint(artboardRect, x, y) {
        return x >= artboardRect[0] && x <= artboardRect[2] && y <= artboardRect[1] && y >= artboardRect[3];
    }

    /**
     * オブジェクト境界[左,上,右,下]とアートボード矩形の重なり面積 / Overlap area between object bounds [L,T,R,B] and an artboard rect
     * @param {number[]} itemBounds - オブジェクトの境界
     * @param {number[]} artboardRect - アートボードの矩形
     * @returns {number} 重なり面積（重ならなければ 0）
     */
    function overlapArea(itemBounds, artboardRect) {
        var overlapWidth = Math.min(itemBounds[2], artboardRect[2]) - Math.max(itemBounds[0], artboardRect[0]);
        var overlapHeight = Math.min(itemBounds[1], artboardRect[1]) - Math.max(itemBounds[3], artboardRect[3]);
        return (overlapWidth > 0 && overlapHeight > 0) ? overlapWidth * overlapHeight : 0;
    }

    /**
     * オブジェクトの所属アートボードを決める（中心点包含→重なり最大→番号が小さい方）
     * Decide which artboard owns an object (center containment → max overlap → smallest index)
     * @param {number[]} itemBounds - オブジェクト境界 [左,上,右,下]
     * @param {number[]} sortedTargets - 昇順の対象アートボード索引
     * @param {object} rectByIndex - index→矩形
     * @returns {number} 所属アートボード索引（どこにも重ならなければ -1）
     */
    function pickOwnerArtboard(itemBounds, sortedTargets, rectByIndex) {
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        /* 1. 中心点が含まれるアートボード（昇順で最初） / center-point containment (first in ascending order) */
        for (var i = 0; i < sortedTargets.length; i++) {
            if (rectContainsPoint(rectByIndex[sortedTargets[i]], centerX, centerY)) { return sortedTargets[i]; }
        }
        /* 2/3. 重なり面積が最大のもの（同じなら昇順の先勝ち＝番号が小さい方） / largest overlap (ties break to the smallest index via ascending scan) */
        var bestIndex = -1;
        var bestArea = 0;
        for (var j = 0; j < sortedTargets.length; j++) {
            var area = overlapArea(itemBounds, rectByIndex[sortedTargets[j]]);
            if (area > bestArea) { bestArea = area; bestIndex = sortedTargets[j]; }
        }
        return bestArea > 0 ? bestIndex : -1;
    }

    /**
     * 対象アートボード群から、重複を排除して所属先ごとにオブジェクトをまとめる
     * Collect objects across target artboards, deduped and grouped by owning artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} sortedTargets - 昇順の対象アートボード索引
     * @param {object} rectByIndex - index→元の矩形
     * @returns {object} index→[所属オブジェクト] / index → owned items
     */
    function collectOwnedObjectsByArtboard(doc, sortedTargets, rectByIndex) {
        var ownedItemsByIndex = {};
        for (var i = 0; i < sortedTargets.length; i++) { ownedItemsByIndex[sortedTargets[i]] = []; }

        /* 重複判定はオブジェクト参照で行う（uuid優先、無ければ === 比較） / Dedupe by object identity (uuid first; === fallback) */
        var seenUuids = {};
        var seenItems = [];

        /**
         * 既に処理したオブジェクトか判定し、未処理なら記録する
         * @param {PageItem} pageItem - 対象のオブジェクト
         * @returns {boolean} 既に処理済みなら true
         */
        function alreadySeen(pageItem) {
            var uuid = getItemUuid(pageItem);
            if (uuid !== null) {
                if (seenUuids[uuid]) { return true; }
                seenUuids[uuid] = true;
                return false;
            }
            for (var k = 0; k < seenItems.length; k++) {
                if (seenItems[k] === pageItem) { return true; } /* 同一参照なら重複 / same reference = duplicate */
            }
            seenItems.push(pageItem);
            return false;
        }

        for (var j = 0; j < sortedTargets.length; j++) {
            var artboardItems = getObjectsOnArtboard(doc, sortedTargets[j]);
            for (var k = 0; k < artboardItems.length; k++) {
                if (alreadySeen(artboardItems[k])) { continue; }
                var ownerIndex = pickOwnerArtboard(artboardItems[k].geometricBounds, sortedTargets, rectByIndex);
                if (ownerIndex !== -1) { ownedItemsByIndex[ownerIndex].push(artboardItems[k]); }
            }
        }
        return ownedItemsByIndex;
    }

    // =========================================
    // 設定の読み取りと寸法計算 / Settings & geometry
    // =========================================

    /**
     * 基準点を固定したまま新しいアートボード矩形を算出する（ピクセルグリッド時は幅高さと基準点を整数化）
     * Compute the new artboard rect keeping the anchor fixed (integerize size & pivot when pixel-grid is on)
     *
     * ピクセルグリッド仕様 / Pixel-grid behavior:
     *   基準点固定と「幅・高さの整数化」を優先する。基準点も整数化するため、
     *   基準点でない側の端（右端／下端など）は 0.5px 単位になる場合がある。
     *   Anchor-fixed and integer width/height take priority; since the pivot is also integerized,
     *   the non-anchored edges (e.g. right/bottom) may land on 0.5px.
     *
     * effectiveRatioX/Y は整数化後の実効倍率（オブジェクト変形はこれを使い、アートボードと一致させる）。
     * effectiveRatioX/Y are the post-rounding effective ratios; object scaling uses them so objects match the artboard.
     * @param {number[]} originalRect - 元のアートボード矩形 [左,上,右,下]
     * @param {object} scaleSettings - readScaleSettings() の戻り値
     * @returns {object} { rect:[左,上,右,下], pivotX, pivotY, effectiveRatioX, effectiveRatioY }
     */
    function computeNewGeometry(originalRect, scaleSettings) {
        var width = originalRect[2] - originalRect[0];
        var height = originalRect[1] - originalRect[3];
        /* 1. 元の矩形から基準点（固定される点）を計算 / pivot (fixed point) from the original rect */
        var pivotX = originalRect[0] + scaleSettings.anchorFx * width;
        var pivotY = originalRect[1] - scaleSettings.anchorFy * height;

        /* 2. 新しい幅・高さ / new width & height */
        var newWidth = width * scaleSettings.ratioX;
        var newHeight = height * scaleSettings.ratioY;

        if (scaleSettings.pixelGrid) {
            /* 3. 幅・高さを整数化（0以下防止） / integerize size (prevent <= 0) */
            newWidth = Math.round(newWidth);
            newHeight = Math.round(newHeight);
            if (newWidth < 1) { newWidth = 1; }
            if (newHeight < 1) { newHeight = 1; }
            /* 4. 基準点も整数化 / integerize the pivot too */
            pivotX = Math.round(pivotX);
            pivotY = Math.round(pivotY);
        }

        /* 5. 基準点を固定して矩形を再計算（Y座標は下方向がマイナス） / rebuild the rect keeping the pivot fixed (Y grows downward negative) */
        var newLeft = pivotX - scaleSettings.anchorFx * newWidth;
        var newTop = pivotY + scaleSettings.anchorFy * newHeight;
        return {
            rect: [newLeft, newTop, newLeft + newWidth, newTop - newHeight],
            pivotX: pivotX,
            pivotY: pivotY,
            /* 整数化後の実効倍率（幅・高さが0のケースは1倍扱い） / post-rounding effective ratio (treat 0 size as 1x) */
            effectiveRatioX: width !== 0 ? newWidth / width : 1,
            effectiveRatioY: height !== 0 ? newHeight / height : 1
        };
    }

    /**
     * ダイアログの対象指定から処理するアートボード索引配列を得る / Resolve target artboard indices
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {number[]|null} 索引配列（不正は null）
     */
    function resolveTargetIndices(resizeDialog, artboardCount) {
        if (resizeDialog.currentRadio.value) {
            /* 現在（起動時）のアクティブアートボードのみ / only the active artboard at launch */
            return [resizeDialog.currentArtboardIndex];
        }
        if (resizeDialog.allRadio.value) {
            var allIndices = [];
            for (var i = 0; i < artboardCount; i++) { allIndices.push(i); }
            return allIndices;
        }
        return parseArtboardSelection(resizeDialog.selectInput.text, artboardCount);
    }

    /**
     * ダイアログの現在値を読み取る / Read the current dialog settings
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {object|null} indices / ratioX / ratioY / anchorFx / anchorFy / scaleObjects / pixelGrid（不正なら null）
     */
    function readScaleSettings(resizeDialog, artboardCount) {
        var targetIndices = resolveTargetIndices(resizeDialog, artboardCount);
        if (!targetIndices) { return null; }
        var scalePercent = parseFloat(resizeDialog.scaleInput.text);
        if (isNaN(scalePercent) || scalePercent <= 0) { return null; }
        var scaleRatio = scalePercent / 100; /* スケールは一意（縦横同率） / Uniform scale (same ratio for both axes) */
        return {
            indices: targetIndices,
            ratioX: scaleRatio,
            ratioY: scaleRatio,
            anchorFx: resizeDialog.anchorX, /* 0=左,0.5=中央,1=右 / 0=left,0.5=center,1=right */
            anchorFy: resizeDialog.anchorY, /* 0=上,0.5=中央,1=下 / 0=top,0.5=center,1=bottom */
            scaleObjects: resizeDialog.scaleObjectsCheckbox.value,
            pixelGrid: resizeDialog.pixelGridCheckbox.value
        };
    }

    // =========================================
    // プレビューと確定 / Preview & commit
    // =========================================

    /**
     * プレビューと確定を行うコントローラを生成する / Create the preview & commit controller
     * プレビューはアートボード矩形のみを更新（完全可逆・線幅に無関係）。オブジェクトはOK確定時に一度だけ resize() する。
     * The live preview updates artboard rectangles only (fully reversible, unrelated to line width);
     * objects are scaled once with resize() at commit. Initial state is the baseline for restore/re-apply to limit drift.
     * オブジェクトは複数アートボードにまたがっても所属ルールで一意化し1回だけ変形する。
     * @param {Document} doc - 対象ドキュメント
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {object} { update, commit, restore, restoreSelectionAndActive, hasError }
     */
    function createPreviewController(doc, resizeDialog, artboardCount) {
        /* 初期状態を保存（キャンセル・確定時に復元） / Save initial state (restored on cancel/commit) */
        var originalActiveIndex = doc.artboards.getActiveArtboardIndex();
        var originalSelection = [];
        var initialSelection = doc.selection;
        if (initialSelection && typeof initialSelection.length === "number") {
            for (var i = 0; i < initialSelection.length; i++) { originalSelection.push(initialSelection[i]); }
        }

        var capturedRects = {};  /* index -> [元rect] / index -> original rect */
        var lastError = null;    /* 直近のエラー / last error */

        /**
         * アートボードの元rectを一度だけ保存する / Capture an artboard's original rect once
         * @param {number} artboardIndex - アートボードの索引
         * @returns {number[]} 元の矩形
         */
        function captureRect(artboardIndex) {
            if (!capturedRects[artboardIndex]) {
                var artboardRect = doc.artboards[artboardIndex].artboardRect;
                capturedRects[artboardIndex] = [artboardRect[0], artboardRect[1], artboardRect[2], artboardRect[3]];
            }
            return capturedRects[artboardIndex];
        }

        /**
         * アートボード矩形を初期スナップショットへ戻す（成功可否を返す） / Reset artboard rects to the initial snapshot (returns success)
         * @returns {object} { success:boolean[, error:Error] }
         */
        function restore() {
            try {
                for (var key in capturedRects) {
                    if (!capturedRects.hasOwnProperty(key)) { continue; }
                    doc.artboards[Number(key)].artboardRect = capturedRects[key];
                }
                return { success: true };
            } catch (e) {
                /* 復元失敗時は履歴（capturedRects）を消さない / keep the history (capturedRects) on failure */
                return { success: false, error: e };
            }
        }

        /**
         * 元の選択状態とアクティブアートボードを可能な範囲で復元する / Restore the original selection and active artboard as far as possible
         * @returns {void}
         */
        function restoreSelectionAndActive() {
            try { doc.selection = null; } catch (e0) {}
            for (var i = 0; i < originalSelection.length; i++) {
                /* 削除・無効化された項目は個別にスキップ / skip deleted/invalid items individually */
                try { originalSelection[i].selected = true; } catch (e1) {}
            }
            try { doc.artboards.setActiveArtboardIndex(originalActiveIndex); } catch (e2) {}
        }

        /**
         * 昇順の対象索引配列を返す（所属ルールのタイブレーク用） / Return target indices sorted ascending (for ownership tie-breaks)
         * @param {object} scaleSettings - readScaleSettings() の戻り値
         * @returns {number[]} 昇順の索引
         */
        function sortedTargetsOf(scaleSettings) {
            var sortedIndices = scaleSettings.indices.concat();
            sortedIndices.sort(function (indexA, indexB) { return indexA - indexB; });
            return sortedIndices;
        }

        /**
         * プレビュー：対象アートボードの矩形だけを更新する（オブジェクトは触らない） / Preview: update only the target artboard rectangles (objects untouched)
         * @returns {void}
         */
        function update() {
            lastError = null;
            try {
                restore(); /* まず矩形を初期状態へ / reset rects to the initial state first */

                var scaleSettings = readScaleSettings(resizeDialog, artboardCount);
                if (!scaleSettings) { app.redraw(); return; } /* 不正入力中は何も適用しない / apply nothing while input is invalid */

                var sortedTargets = sortedTargetsOf(scaleSettings);
                for (var i = 0; i < sortedTargets.length; i++) {
                    var artboardIndex = sortedTargets[i];
                    captureRect(artboardIndex);
                    doc.artboards[artboardIndex].artboardRect = computeNewGeometry(capturedRects[artboardIndex], scaleSettings).rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            }
            app.redraw();
        }

        /**
         * 確定：矩形を初期状態へ戻し、初期状態から一度だけ「矩形＋オブジェクト」を適用する（線幅も正しく拡縮・累積誤差なし）
         * Commit: reset rects, then apply "rects + objects" once from the initial state (correct line width, no drift)
         * @returns {void}
         */
        function commit() {
            lastError = null;
            /* position/geometricBounds を artboardRect と同じドキュメント座標で扱うため明示設定（終了時に復元） / Force document coordinates so position/geometricBounds match artboardRect (restored at the end) */
            var previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
            try {
                restore(); /* プレビューの矩形変更を初期状態へ戻す / reset the preview rects to the initial state */

                var scaleSettings = readScaleSettings(resizeDialog, artboardCount);
                if (!scaleSettings) { app.redraw(); return; }

                var sortedTargets = sortedTargetsOf(scaleSettings);
                var i;
                for (i = 0; i < sortedTargets.length; i++) { captureRect(sortedTargets[i]); }
                var ownedObjects = scaleSettings.scaleObjects
                    ? collectOwnedObjectsByArtboard(doc, sortedTargets, capturedRects)
                    : {};

                var committedResizes = []; /* 失敗時のベストエフォート巻き戻し用 / for best-effort rollback on failure */
                for (i = 0; i < sortedTargets.length; i++) {
                    var artboardIndex = sortedTargets[i];
                    var geometry = computeNewGeometry(capturedRects[artboardIndex], scaleSettings);

                    /* オブジェクトはアートボードと同じ基準点・実効倍率（整数化後）で拡縮 / Objects use the artboard's pivot and post-rounding effective ratio */
                    if (scaleSettings.scaleObjects) {
                        var ownedItems = ownedObjects[artboardIndex] || [];
                        var resizeResult = resizeItemsAbout(ownedItems, geometry.effectiveRatioX, geometry.effectiveRatioY, geometry.pivotX, geometry.pivotY);
                        committedResizes.push({ items: resizeResult.transformedItems, ratioX: geometry.effectiveRatioX, ratioY: geometry.effectiveRatioY, pivotX: geometry.pivotX, pivotY: geometry.pivotY });
                        if (!resizeResult.success) {
                            /* 途中失敗：確定済みを逆拡縮し、矩形を戻して中止 / mid-way failure: inverse-resize committed items, reset rects, and stop */
                            lastError = resizeResult.error;
                            for (var j = committedResizes.length - 1; j >= 0; j--) {
                                resizeItemsAbout(committedResizes[j].items, 1 / committedResizes[j].ratioX, 1 / committedResizes[j].ratioY, committedResizes[j].pivotX, committedResizes[j].pivotY);
                            }
                            restore();
                            app.redraw();
                            return;
                        }
                    }

                    doc.artboards[artboardIndex].artboardRect = geometry.rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            } finally {
                /* 座標系を元へ戻す / restore the coordinate system */
                try { app.coordinateSystem = previousCoordinateSystem; } catch (e2) {}
            }
            app.redraw();
        }

        return {
            update: update,
            commit: commit,
            restore: restore,
            restoreSelectionAndActive: restoreSelectionAndActive,
            hasError: function () { return lastError !== null; }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * キャンセル・失敗時の後始末：矩形と選択・アクティブアートボードを戻して再描画する
     * @param {object} previewController - createPreviewController() の戻り値
     * @returns {object} 矩形の復元結果 { success:boolean[, error:Error] }
     */
    function revertAll(previewController) {
        var restoreResult = previewController.restore();
        previewController.restoreSelectionAndActive();
        app.redraw();
        return restoreResult;
    }

    /**
     * エントリポイント：ダイアログを出し、アートボードサイズをプレビューしつつ確定でオブジェクトも拡縮する
     * Entry point: show dialog, preview artboard size, scale objects on commit
     * @returns {void}
     */
    function resizeArtboards() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var rulerUnit = getUnitInfo();
        var artboardCount = doc.artboards.length;

        /* アクティブアートボードの現在サイズを初期値にする / Use the active artboard's current size as defaults */
        var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var defaultWidth = formatNumber((activeRect[2] - activeRect[0]) / rulerUnit.pointsPerUnit);
        var defaultHeight = formatNumber((activeRect[1] - activeRect[3]) / rulerUnit.pointsPerUnit);

        var resizeDialog = createResizeDialog(rulerUnit.label, defaultWidth, defaultHeight, artboardCount);
        /* 「現在のアートボード」対象用に起動時のアクティブ索引を保持 / Remember the launch-time active index for the "Current artboard" target */
        resizeDialog.currentArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* プレビュー配線（矩形のみ）＋表示時にスケール欄へフォーカス / Wire up the preview (rects only) and focus the scale field on show */
        var previewController = createPreviewController(doc, resizeDialog, artboardCount);
        resizeDialog.onPreview = previewController.update;
        resizeDialog.onShow = function () {
            previewController.update();
            resizeDialog.scaleInput.active = true;
        };

        prepareDialogWindow(resizeDialog, SCRIPT_NAME);
        if (resizeDialog.show() !== 1) {
            /* キャンセル：矩形を復元（失敗は通知）し、選択・アクティブアートボードを戻す / Cancel: restore rects (notify on failure), restore selection & active artboard */
            if (!revertAll(previewController).success) { alert(getLabel("alert.restoreError")); }
            return;
        }

        /* OK：不正値は矩形を戻してエラー表示 / OK: revert rects and report on invalid input */
        if (!readScaleSettings(resizeDialog, artboardCount)) {
            revertAll(previewController);
            alert(resolveTargetIndices(resizeDialog, artboardCount) ? getLabel("alert.invalidNumber") : getLabel("alert.invalidSelection"));
            return;
        }

        /* 確定：矩形を戻し、resize()で「矩形＋オブジェクト」を一度だけ適用 / Commit: reset rects, then apply rects + objects once via resize() */
        previewController.commit();
        if (previewController.hasError()) {
            /* 変形失敗時は確定せず、矩形と選択・アクティブを復元 / on failure, do not commit; restore rects, selection & active */
            revertAll(previewController);
            alert(getLabel("alert.transformError"));
            return;
        }

        /* 適用済みの状態を確定。選択・アクティブアートボードのみ復元（完了メッセージは表示しない） / Keep the applied result; restore only selection & active artboard (no completion message) */
        previewController.restoreSelectionAndActive();
        app.redraw();
    }

    resizeArtboards();

})();
