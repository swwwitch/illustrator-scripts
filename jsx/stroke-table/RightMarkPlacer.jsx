#target illustrator
#targetengine "RightMarkPlacerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクトを左から順に見て、隣り合うオブジェクト同士のアキの中央に記号を配置します。
記号は9種類から選べ、高さ・幅・線幅・位置をダイアログで調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nebac730ec187

### Overview

Scans the selected objects from left to right and places a mark in the middle of the gap between each adjacent pair.
Nine marks are available, with height, width, stroke weight and position set from the dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RightMarkPlacer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RightMarkPlacer";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RightMarkPlacer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nebac730ec187"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 記号の色（CMYK）/ Mark color (CMYK) */
    var MARK_COLOR_CMYK = [0, 0, 0, 100];

    /* 高さ（％）の上限 / Maximum height percentage */
    var MAX_HEIGHT_PERCENT = 200;

    /* 線幅の下限（pt）/ Minimum stroke width in points */
    var MIN_STROKE_WIDTH_PT = 0.25;

    /* ▶ の凹みの上限（幅に対する割合）/ Maximum inset as a ratio of the width */
    var MAX_INSET_RATIO = 0.8;

    /* ▶ の角丸半径（幅と高さの小さい方に対する割合）/ Rounded-corner radius as a ratio of the smaller side */
    var TRIANGLE_CORNER_RADIUS_RATIO = 0.12;

    /* ＿\ の斜線の角度（度）/ Slash angle in degrees */
    var SLASH_ANGLE_DEFAULT = 35;
    var SLASH_ANGLE_MAX = 89;

    /* ➡ の矢じりの天地を求める、線幅に対する倍率 / Height of the solid arrowhead as a multiple of the stroke width */
    var SOLID_ARROW_HEIGHT_TO_STROKE_RATIO = 3;

    /* → ➡ の矢じりの奥行きが、幅に占める割合の上限 / Maximum arrowhead depth as a ratio of the width */
    var MAX_ARROW_HEAD_RATIO = 0.9;

    /* 幅を自動計算するときの、アキに対する割合 / Ratio of the gap used when the width is calculated automatically */
    var AUTO_WIDTH_GAP_RATIO = 0.7;
    var AUTO_WIDTH_GAP_RATIO_SMALL = 0.35;

    /* 山形の幅を自動計算するときの、高さに対する割合 / Ratio of the height used when a chevron width is calculated automatically */
    var AUTO_WIDTH_CHEVRON_HEIGHT_RATIO = 0.5;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var RADIO_SPACING  = 6;                  /* ラジオボタンの間隔 / gap between radio buttons */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var COLUMN_PANEL_SPACING = 10;           /* カラム内のパネルの間隔 / gap between panels in a column */
    var FIELD_ROW_SPACING = 8;               /* ラベルと入力欄の間隔 / gap between a label and its field */
    var LABEL_COLUMN_WIDTH = 60;             /* ラベル列の幅 / width of the label column */
    var FIELD_CHARACTERS = 4;                /* 入力欄の幅（文字数）/ field width in characters */
    var MIRROR_ROW_TOP_MARGIN = 6;           /* ［左右逆］の上余白 / top margin above the mirror checkbox */
    var OPTION_CHECKBOX_INDENT = 40;         /* ［オプション］のチェックボックスの字下げ / indent of the option checkboxes */
    var OPTION_CHECKBOX_TOP_MARGIN = 10;     /* ［オプション］のチェックボックスの上余白 / top margin above the option checkboxes */

    /**
     * ウィンドウに共通のレイアウトを適用します。
     *
     * @param {Window} targetWindow - 対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用します。
     *
     * @param {Panel} targetPanel - 対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
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
     * 縦並びのグループ（カラム）に共通のレイアウトを適用します。
     *
     * @param {Group} columnGroup - 対象のグループ。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupColumn(columnGroup, spacing) {
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びのグループ（入力行など）に共通のレイアウトを適用します。
     *
     * @param {Group} rowGroup - 対象のグループ。
     * @param {string} [alignment] - グループ自体の配置。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    var rulerUnitInfo = getUnitInfo("rulerType");
    var strokeUnitInfo = getUnitInfo("strokeUnits");

    /**
     * 単位の数値を pt に換算します。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} pt 値。
     */
    function convertValueToPt(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * pt を単位の数値に換算します。
     *
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} 単位付きの数値。
     */
    function convertPtToUnitValue(valuePt, unitInfo) {
        return valuePt / unitInfo.pointsPerUnit;
    }

    /* 表示桁数 / Decimal places used for display */
    var DISPLAY_DECIMALS = 2;          /* 既定の桁数 / default decimal places */
    var DISPLAY_DECIMALS_MAX = 5;      /* 桁数の上限 / maximum decimal places */
    var DISPLAY_DECIMALS_PT_FACTOR = 3; /* この換算係数までは既定の桁数 / units up to this factor keep the default */

    /**
     * pt 値を小数点以下2桁へ丸めます。
     *
     * @param {number} value - 丸める値。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValue(value) {
        return Math.round(value * 100) / 100;
    }

    /**
     * 単位に応じた表示桁数を求めます。
     * inch のように 1単位が大きい単位では、2桁だと pt 換算で精度が足りないため桁数を増やします。
     *
     * @param {object} [unitInfo] - 単位情報。
     * @returns {number} 小数点以下の桁数。
     */
    function getDisplayDecimals(unitInfo) {
        if (!unitInfo || !(unitInfo.pointsPerUnit > DISPLAY_DECIMALS_PT_FACTOR)) return DISPLAY_DECIMALS;
        var extraDigits = Math.ceil(Math.log(unitInfo.pointsPerUnit) / Math.LN10);
        return Math.min(DISPLAY_DECIMALS + extraDigits, DISPLAY_DECIMALS_MAX);
    }

    /**
     * 単位に合わせた桁数で丸めます。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} [unitInfo] - 単位情報。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValueForUnit(value, unitInfo) {
        var scale = Math.pow(10, getDisplayDecimals(unitInfo));
        return Math.round(value * scale) / scale;
    }

    /**
     * 矢印キーで増減する量を、単位に合わせて求めます。
     * pt・px・mm・Q・H は 1単位、inch のように 1単位が大きい単位はより細かい刻みにします。
     *
     * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
     * @returns {number} 増減量。
     */
    function getArrowKeyStep(unitInfo) {
        if (!unitInfo || !(unitInfo.pointsPerUnit > 0)) return 1;
        var exponent = Math.max(0, Math.round(Math.log(unitInfo.pointsPerUnit) / Math.LN10));
        return Math.pow(10, -exponent);
    }

    /**
     * pt 値を単位に換算して入力欄へ表示します。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {void}
     */
    function setFieldFromPt(editText, valuePt, unitInfo) {
        editText.text = String(roundDisplayValueForUnit(convertPtToUnitValue(valuePt, unitInfo), unitInfo));
    }

    /**
     * 入力欄の値を pt として読み取ります。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {object} unitInfo - 単位情報。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @returns {number} pt 値。数値でない場合と、許可していない負の値の場合は NaN。
     */
    function parseFieldToPt(editText, unitInfo, allowNegative) {
        var value = parseFloat(editText.text);
        if (isNaN(value)) return NaN;
        /* 負の値を黙って 0 にせず、呼び出し元でエラーとして扱えるようにする / Report it instead of silently clamping to zero */
        if (!allowNegative && value < 0) return NaN;
        return convertValueToPt(value, unitInfo);
    }

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

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

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

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            /* ＋ や × のような無方向の形状と［左右逆］があるため、向きを含めない名前にしています / Kept direction-neutral: some shapes have no direction and the mark can be mirrored */
            title: { ja: "オブジェクト間に記号を配置", en: "Place Marks Between Objects" }
        },
        panel: {
            shape: { ja: "形状", en: "Shape" },
            capStyle: { ja: "先端", en: "End Style" },
            options: { ja: "オプション", en: "Options" },
            position: { ja: "位置調整", en: "Position" }
        },
        radio: {
            arrowSlash: { ja: "＿\\", en: "─\\" },
            capNone: { ja: "なし", en: "None" },
            capRound: { ja: "丸型", en: "Round" }
        },
        checkbox: {
            mirrorHorizontal: { ja: "左右逆", en: "Mirror horizontally" },
            flatChevron: { ja: "天地を水平に", en: "Keep top and bottom edges horizontal" },
            roundCorners: { ja: "角丸", en: "Rounded corners" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        /* 入力欄の前に置く項目名。区切りのコロンは labelText() で付ける / Field captions; labelText() adds the colon */
        fieldLabel: {
            height: { ja: "高さ", en: "Height" },
            width: { ja: "幅", en: "Width" },
            gap: { ja: "間隔", en: "Gap" },
            inset: { ja: "凹み", en: "Inset" },
            strokeWidth: { ja: "線幅", en: "Stroke" },
            angle: { ja: "角度", en: "Angle" },
            offsetX: { ja: "左右", en: "Horizontal" },
            offsetY: { ja: "上下", en: "Vertical" }
        },
        tooltip: {
            solidArrow: { ja: "矢じりの天地は線幅の%1倍になります。", en: "The arrowhead is %1× the stroke width tall." },
            capNone: { ja: "ショートカット：F", en: "Shortcut: F" },
            capRound: { ja: "ショートカット：R", en: "Shortcut: R" },
            mirrorHorizontal: { ja: "ショートカット：V", en: "Shortcut: V" },
            height: {
                ja: "隣り合う2つのオブジェクト全体の高さに対する割合（%1%まで）",
                en: "Percentage of the overall height of the two adjacent objects (up to %1%)"
            },
            width: { ja: "空欄にすると自動計算に戻ります", en: "Clear to calculate automatically again" },
            inset: { ja: "幅の%1%が上限です", en: "Max %1% of width" },
            gap: {
                ja: ">> の2つの山形の間隔。負の値で重なります",
                en: "Space between the two chevrons of >>. Negative values overlap them"
            },
            angle: { ja: "＿\\ の斜線の角度（%1°まで）", en: "Slash angle of ─\\ (up to %1°)" },
            flatChevron: {
                ja: "> と >> を、天地が水平な塗りの形状で作成します",
                en: "Draws > and >> as filled shapes with horizontal top and bottom edges"
            },
            roundCorners: {
                ja: "▶ の角を丸くします。半径は記号の大きさから自動で決まります",
                en: "Rounds the corners of ▶. The radius is set automatically from the mark size"
            },
            offsetX: { ja: "記号を左右にずらします。正の値で右へ移動します", en: "Shifts the marks horizontally. Positive values move them right" },
            offsetY: { ja: "記号を上下にずらします。正の値で上へ移動します", en: "Shifts the marks vertically. Positive values move them up" },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTwoObjects: { ja: "オブジェクトを2つ以上選択してください。", en: "Please select two or more objects." },
            lockedLayer: {
                ja: "作業レイヤーがロックまたは非表示です。解除してから実行してください。",
                en: "The active layer is locked or hidden. Please unlock and show it, then run again."
            },
            positiveNumber: { ja: "正の数値を入力してください。", en: "Please enter a positive value." },
            maxHeight: { ja: "%1% 以下の値を入力してください。", en: "Please enter a value of %1% or less." },
            noGap: {
                ja: "隣り合うオブジェクト間に作成できるアキがありません。",
                en: "There is no usable gap between adjacent objects."
            },
            strokePositive: {
                ja: "線幅は %1 %2 以上の値を入力してください。",
                en: "Please enter a stroke width of %1 %2 or greater."
            },
            widthPositive: { ja: "幅は 0 以上の値を入力してください。", en: "Please enter a width value of 0 or greater." },
            invalidValue: { ja: "入力値を確認してください。", en: "Please check the input values." }
        },
        log: {
            measureTextBounds: { ja: "テキストの計測用アウトライン化", en: "Outline text for measurement" },
            removeMeasurementCopy: { ja: "計測用複製の削除", en: "Remove measurement copy" },
            applyRoundCorners: { ja: "角丸効果の適用", en: "Apply rounded-corners effect" },
            restoreSelection: { ja: "選択状態の復元", en: "Restore selection" },
            removePreviewItem: { ja: "プレビューの削除", en: "Remove preview item" },
            mergeSolidArrow: { ja: "➡ の合成", en: "Merge the solid arrow" }
        }
    };

    // =========================================
    // 汎用ユーティリティ / Generic utilities
    // =========================================

    /**
     * コンテキスト付きでエラーを $.writeln に出力します。
     *
     * @param {string} logContext - どの処理で起きたかを示す文言。
     * @param {object} errorObject - エラーオブジェクトまたはメッセージ。
     * @returns {void}
     */
    function logScriptError(logContext, errorObject) {
        $.writeln("[" + SCRIPT_NAME + " " + SCRIPT_VERSION + "] " + logContext + ": " + errorObject);
    }

    /**
     * 処理を実行し、失敗した場合はログを出して続行します。
     *
     * @param {Function} operation - 実行する処理。
     * @param {string} logContext - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true、失敗したら false。
     */
    function runSafely(operation, logContext) {
        try {
            operation();
            return true;
        } catch (e) {
            logScriptError(logContext, e);
            return false;
        }
    }

    /**
     * ExtendScript のコレクションを通常の配列へコピーします。
     *
     * @param {object} collection - selection などのコレクション。
     * @returns {Array} コピーした配列。
     */
    function collectionToArray(collection) {
        var copiedItems = [];
        if (!collection) return copiedItems;
        for (var i = 0; i < collection.length; i++) {
            copiedItems.push(collection[i]);
        }
        return copiedItems;
    }

    /**
     * アイテムを削除します。削除済みのアイテムを渡しても止まらないよう、失敗はログだけにします。
     *
     * @param {PageItem} pageItem - 対象のアイテム。null なら何もしません。
     * @param {string} logContext - 失敗時にログへ出す文言。
     * @returns {void}
     */
    function removeItemSafely(pageItem, logContext) {
        if (!pageItem) return;
        runSafely(function () {
            pageItem.remove();
        }, logContext);
    }

    /**
     * 選択状態をまとめて復元します。
     * メニューコマンドで置き換わったアイテムが混ざっても止まらないよう、1つずつ選択し直します。
     *
     * @param {PageItem[]} pageItems - 選択し直すアイテムの配列。
     * @returns {void}
     */
    function restoreSelection(pageItems) {
        app.activeDocument.selection = null;
        for (var i = 0; i < pageItems.length; i++) {
            (function (pageItem) {
                runSafely(function () {
                    pageItem.selected = true;
                }, getLabel("log.restoreSelection"));
            })(pageItems[i]);
        }
    }

    // =========================================
    // 選択オブジェクトの計測 / Measuring the selection
    // =========================================

    /*
     * 同じアイテムを何度も計測しないためのキャッシュ。
     * モーダルダイアログの表示中は選択オブジェクトが変化しないため、閉じるまで保持します。
     * テキストの計測はアウトライン化を伴うので、プレビューの更新ごとに測り直すと重くなります。
     *
     * Cache so each item is measured only once.
     * The selection cannot change while the modal dialog is up, so it is kept until the dialog closes.
     * Measuring text requires outlining it, which is too slow to repeat on every preview refresh.
     */
    var measuredBoundsCache = [];

    /**
     * テキストをアウトライン化した状態のバウンズを求めます。
     * 複製をアウトライン化して計測し、計測用のアイテムは必ず削除します。
     *
     * @param {TextFrame} textFrame - 対象のテキスト。
     * @returns {number[]} バウンズ。失敗した場合は元のバウンズ。
     */
    function measureOutlinedTextBounds(textFrame) {
        var duplicatedText = null;
        var outlinedText = null;
        try {
            duplicatedText = textFrame.duplicate();
            outlinedText = duplicatedText.createOutline();
            return outlinedText.geometricBounds;
        } catch (e) {
            logScriptError(getLabel("log.measureTextBounds"), e);
            return textFrame.geometricBounds;
        } finally {
            /* createOutline() は複製を消費するので、アウトラインができたらそちらだけを消す / createOutline() consumes the duplicate, so remove the outline once it exists */
            removeItemSafely(outlinedText || duplicatedText, getLabel("log.removeMeasurementCopy"));
        }
    }

    /**
     * 計測用のバウンズを取得します（テキストはアウトライン化した形状で計測）。
     * 戻り値はキャッシュそのものなので、呼び出し側では書き換えません。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @returns {number[]} バウンズ。
     */
    function getItemMeasurementBounds(pageItem) {
        for (var i = 0; i < measuredBoundsCache.length; i++) {
            if (measuredBoundsCache[i].item === pageItem) return measuredBoundsCache[i].bounds;
        }

        var measuredBounds = (pageItem.typename === "TextFrame")
            ? measureOutlinedTextBounds(pageItem)
            : pageItem.geometricBounds;

        measuredBoundsCache.push({ item: pageItem, bounds: measuredBounds });
        return measuredBounds;
    }

    /**
     * アイテムの中心の X 座標を求めます。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @returns {number} 中心の X 座標。
     */
    function getItemCenterX(pageItem) {
        var itemBounds = getItemMeasurementBounds(pageItem);
        return (itemBounds[0] + itemBounds[2]) / 2;
    }

    /**
     * アイテムの配列を、中心の X 座標で左から右の順に並べ替えます（配列そのものを並べ替えます）。
     *
     * @param {PageItem[]} pageItems - 対象のアイテムの配列。
     * @returns {PageItem[]} 並べ替えた配列。
     */
    function sortItemsLeftToRight(pageItems) {
        return pageItems.sort(function (leftItem, rightItem) {
            return getItemCenterX(leftItem) - getItemCenterX(rightItem);
        });
    }

    /**
     * 隣り合う2つのオブジェクトから、記号を置く位置とサイズを求めます。
     *
     * @param {PageItem} leftItem - 左側のアイテム。
     * @param {PageItem} rightItem - 右側のアイテム。
     * @param {number} heightPercent - 高さ（％）。null なら合計高さをそのまま使います。
     * @param {number} widthPt - 幅（pt）。0以下なら高さから決めます。
     * @param {number} offsetX - 左右の位置調整（pt）。
     * @param {number} offsetY - 上下の位置調整（pt）。
     * @returns {object} 配置情報。アキがない場合は null。
     */
    function computePlacementBetweenItems(leftItem, rightItem, heightPercent, widthPt, offsetX, offsetY) {
        var firstBounds = getItemMeasurementBounds(leftItem);
        var secondBounds = getItemMeasurementBounds(rightItem);

        var leftBounds = (firstBounds[0] < secondBounds[0]) ? firstBounds : secondBounds;
        var rightBounds = (firstBounds[0] < secondBounds[0]) ? secondBounds : firstBounds;

        var gapLeft = leftBounds[2];
        var gapRight = rightBounds[0];
        if (gapRight - gapLeft <= 0) return null;

        var top = Math.max(firstBounds[1], secondBounds[1]);
        var bottom = Math.min(firstBounds[3], secondBounds[3]);
        var totalHeight = top - bottom;
        var usesFullHeight = (heightPercent === null || typeof heightPercent === "undefined");
        var markHeight = usesFullHeight ? totalHeight : (totalHeight * (heightPercent / 100));

        return {
            gapLeft: gapLeft,
            gapRight: gapRight,
            centerX: (gapLeft + gapRight) / 2 + (offsetX || 0),
            centerY: (top + bottom) / 2 + (offsetY || 0),
            totalHeight: totalHeight,
            height: markHeight,
            width: (widthPt > 0) ? widthPt : markHeight,
            /* 矢印系は幅未指定のとき、アキに対する割合で決める / Arrow shapes fall back to a ratio of the gap when no width is set */
            arrowWidth: (widthPt > 0) ? widthPt : ((gapRight - gapLeft) * AUTO_WIDTH_GAP_RATIO)
        };
    }

    /**
     * 隣り合う組を走査し、最も狭いアキと最も低い合計高さを求めます。
     *
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @returns {object} minGapWidth / minTotalHeight を持つオブジェクト。アキのある組がなければ null。
     */
    function measureNarrowestGap(sortedItems) {
        var minGapWidth = null;
        var minTotalHeight = null;

        for (var i = 0; i < sortedItems.length - 1; i++) {
            var placement = computePlacementBetweenItems(sortedItems[i], sortedItems[i + 1], null, 0, 0, 0);
            if (!placement) continue;

            var gapWidth = placement.gapRight - placement.gapLeft;
            if (minGapWidth === null || gapWidth < minGapWidth) {
                minGapWidth = gapWidth;
            }
            if (minTotalHeight === null || placement.totalHeight < minTotalHeight) {
                minTotalHeight = placement.totalHeight;
            }
        }

        if (minGapWidth === null) return null;
        return { minGapWidth: minGapWidth, minTotalHeight: minTotalHeight };
    }

    // =========================================
    // 形状定義 / Shape definitions
    // =========================================

    /* 形状の設定の既定値。各形状には、既定と違う項目だけを書きます / Shared defaults; each shape lists only what differs */
    var SHAPE_CONFIG_DEFAULTS = {
        symbol: "",                   /* ラジオボタンの表記 / radio button label */
        helpTip: "",                  /* ラジオボタンの tooltip / radio button tooltip */
        enableCapPanel: true,         /* ［先端］パネル / End Style panel */
        enableStrokeInput: true,      /* ［線幅］（有効なら下限も検証）/ stroke width, validated against the minimum */
        enableHeightInput: false,     /* ［高さ］/ height */
        enableMirror: false,          /* ［左右逆］/ mirror */
        enableGap: false,             /* ［間隔］/ gap */
        enableInset: false,           /* ［凹み］/ inset */
        enableRoundCorners: false,    /* ［角丸］/ rounded corners */
        enableAngle: false,           /* ［角度］/ angle */
        enableFlatChevron: false,     /* ［天地を水平に］/ flat chevron */
        changesSelection: false,      /* 作成時にメニューコマンドで選択を変えるか / whether creation drives menu commands on the selection */
        defaultHeightPercent: 0,
        defaultGap: -1,
        defaultStrokePt: 0.6,
        defaultAngle: 0,
        /* 幅の自動計算：basis が "gap" なら最も狭いアキ、"height" なら高さ（％）を掛けた合計高さに ratio を掛ける
           Automatic width: ratio × the narrowest gap ("gap"), or × the total height scaled by the height percentage ("height") */
        autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO },
        createMark: null              /* 記号を作成する関数 / function that draws the mark */
    };

    /**
     * 既定値に、形状ごとの設定を重ねたオブジェクトを返します。
     *
     * @param {object} overrides - 既定値と違う項目。
     * @returns {object} 形状の設定。
     */
    function withShapeDefaults(overrides) {
        var shapeConfig = {};
        var settingKey;
        for (settingKey in SHAPE_CONFIG_DEFAULTS) {
            if (SHAPE_CONFIG_DEFAULTS.hasOwnProperty(settingKey)) shapeConfig[settingKey] = SHAPE_CONFIG_DEFAULTS[settingKey];
        }
        for (settingKey in overrides) {
            if (overrides.hasOwnProperty(settingKey)) shapeConfig[settingKey] = overrides[settingKey];
        }
        return shapeConfig;
    }

    /* 形状ごとの設定 / Per-shape settings */
    var SHAPE_CONFIG = {
        triangle: withShapeDefaults({
            symbol: "▶",
            enableCapPanel: false,
            /* 塗りだけで描くので線幅は使いません / Drawn as a fill only, so the stroke width is unused */
            enableStrokeInput: false,
            enableHeightInput: true,
            enableMirror: true,
            enableInset: true,
            enableRoundCorners: true,
            defaultHeightPercent: 30,
            defaultStrokePt: 0.3,
            autoWidth: { basis: "height", ratio: 1 },
            createMark: createTriangleMark
        }),
        chevron: withShapeDefaults({
            symbol: ">",
            enableHeightInput: true,
            enableMirror: true,
            enableFlatChevron: true,
            defaultHeightPercent: 50,
            autoWidth: { basis: "height", ratio: AUTO_WIDTH_CHEVRON_HEIGHT_RATIO },
            createMark: createChevronMark
        }),
        doubleChevron: withShapeDefaults({
            symbol: ">>",
            enableHeightInput: true,
            enableMirror: true,
            enableGap: true,
            enableFlatChevron: true,
            defaultHeightPercent: 50,
            defaultGap: 0,
            /* 幅は山形1つ分。全体の幅は「幅×2＋間隔」になります / The width is per chevron; the pair spans width × 2 + gap */
            autoWidth: { basis: "height", ratio: AUTO_WIDTH_CHEVRON_HEIGHT_RATIO },
            createMark: createDoubleChevronMark
        }),
        dash: withShapeDefaults({
            symbol: "─",
            createMark: createDashMark
        }),
        arrow: withShapeDefaults({
            symbol: "→",
            enableHeightInput: true,
            enableMirror: true,
            defaultHeightPercent: 30,
            createMark: createArrowMark
        }),
        solidArrow: withShapeDefaults({
            symbol: "➡",
            helpTip: getLabel("tooltip.solidArrow", [SOLID_ARROW_HEIGHT_TO_STROKE_RATIO]),
            /* 天地は線幅から決まるため、高さ（％）は使いません（enableHeightInput は既定の false）/ The height comes from the stroke width, so the percentage stays disabled */
            enableCapPanel: false,
            enableMirror: true,
            changesSelection: true,
            defaultStrokePt: 3.6,
            createMark: createSolidArrowMark
        }),
        arrowSlash: withShapeDefaults({
            symbol: getLabel("radio.arrowSlash"),
            enableHeightInput: true,
            enableMirror: true,
            enableAngle: true,
            defaultHeightPercent: 50,
            defaultAngle: SLASH_ANGLE_DEFAULT,
            createMark: createArrowSlashMark
        }),
        plus: withShapeDefaults({
            symbol: "＋",
            autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO_SMALL },
            createMark: createPlusMark
        }),
        multiply: withShapeDefaults({
            symbol: "×",
            autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO_SMALL },
            createMark: createMultiplyMark
        })
    };

    /* ［形状］パネルに並べる順。先頭が初期選択 / Order in the Shape panel; the first one starts selected */
    var SHAPE_ORDER = ["triangle", "chevron", "doubleChevron", "dash", "arrow", "solidArrow", "arrowSlash", "plus", "multiply"];

    // =========================================
    // パス作成のヘルパー / Path creation helpers
    // =========================================

    /**
     * 記号を作成するレイヤー（作業レイヤー）を返します。
     *
     * @returns {Layer} 作業レイヤー。
     */
    function getTargetLayer() {
        return app.activeDocument.activeLayer;
    }

    /**
     * 記号の色を作成します。
     *
     * @returns {CMYKColor} 記号の色。
     */
    function createMarkColor() {
        var markColor = new CMYKColor();
        markColor.cyan = MARK_COLOR_CMYK[0];
        markColor.magenta = MARK_COLOR_CMYK[1];
        markColor.yellow = MARK_COLOR_CMYK[2];
        markColor.black = MARK_COLOR_CMYK[3];
        return markColor;
    }

    /**
     * 線端（と角の形状）を設定します。
     *
     * @param {PathItem} pathItem - 対象のパス。
     * @param {object} endStyle - round（丸型にするか）/ join（角の形状も設定するか）/ flatWhenNotRound（丸型でないとき線端なし・マイターを明示するか）。
     * @returns {void}
     */
    function applyStrokeEnds(pathItem, endStyle) {
        if (endStyle.round) {
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            if (endStyle.join) pathItem.strokeJoin = StrokeJoin.ROUNDENDJOIN;
            return;
        }
        if (endStyle.flatWhenNotRound) {
            pathItem.strokeCap = StrokeCap.BUTTENDCAP;
            if (endStyle.join) pathItem.strokeJoin = StrokeJoin.MITERENDJOIN;
        }
    }

    /**
     * 線だけの開いたパスを作成します。
     *
     * @param {number[][]} points - アンカーポイントの配列。
     * @param {number} strokeWidthPt - 線幅（pt）。
     * @param {object} [endStyle] - applyStrokeEnds に渡す線端の指定。省略時は線端を変えません。
     * @returns {PathItem} 作成したパス。
     */
    function createStrokedPath(points, strokeWidthPt, endStyle) {
        var strokedPath = getTargetLayer().pathItems.add();
        strokedPath.setEntirePath(points);
        strokedPath.closed = false;
        strokedPath.filled = false;
        strokedPath.stroked = true;
        strokedPath.strokeWidth = strokeWidthPt;
        strokedPath.strokeColor = createMarkColor();
        if (endStyle) applyStrokeEnds(strokedPath, endStyle);
        return strokedPath;
    }

    /**
     * 2点を結ぶ直線を、入力値の線幅と線端で作成します。
     *
     * @param {number[]} startPoint - 始点。
     * @param {number[]} endPoint - 終点。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createStrokedLine(startPoint, endPoint, markSettings) {
        return createStrokedPath([startPoint, endPoint], markSettings.strokeWidthPt, { round: markSettings.roundEnds });
    }

    /**
     * 塗りだけの閉じたパスを作成します。
     *
     * @param {number[][]} points - アンカーポイントの配列。
     * @returns {PathItem} 作成したパス。
     */
    function createFilledPath(points) {
        var filledPath = getTargetLayer().pathItems.add();
        filledPath.setEntirePath(points);
        filledPath.closed = true;
        filledPath.filled = true;
        filledPath.fillColor = createMarkColor();
        filledPath.stroked = false;
        return filledPath;
    }

    /**
     * 複数のアイテムを1つのグループにまとめます。
     *
     * @param {PageItem[]} pageItems - まとめるアイテムの配列。
     * @returns {GroupItem} 作成したグループ。
     */
    function groupPageItems(pageItems) {
        var markGroup = getTargetLayer().groupItems.add();
        for (var i = 0; i < pageItems.length; i++) {
            pageItems[i].move(markGroup, ElementPlacement.INSIDE);
        }
        return markGroup;
    }

    /**
     * 角丸のライブエフェクトを適用します。
     * 効果はプラグイン側で解釈されるため、適用に失敗しても角丸なしで続行します。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @param {number} radiusPt - 角丸の半径（pt）。
     * @returns {void}
     */
    function applyRoundCornersEffect(pageItem, radiusPt) {
        runSafely(function () {
            pageItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
        }, getLabel("log.applyRoundCorners"));
    }

    // =========================================
    // 形状の作成 / Shape creation
    // =========================================
    // 各 create…Mark() は、配置情報（computePlacementBetweenItems の戻り値）と
    // 検証済みの入力値（readMarkSettings の戻り値）を受け取ります。
    // Each create…Mark() takes the placement (from computePlacementBetweenItems)
    // and the validated settings (from readMarkSettings).

    /**
     * ▶ の角丸半径を求めます。
     *
     * @param {number} width - 幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @param {number} insetPt - 凹み（pt）。
     * @returns {number} 角丸の半径（pt）。
     */
    function getTriangleCornerRadius(width, height, insetPt) {
        var radius = Math.min(width, height) * TRIANGLE_CORNER_RADIUS_RATIO;
        if (insetPt > 0) {
            radius = Math.min(radius, insetPt * 0.45);
        }
        return Math.max(1, radius);
    }

    /**
     * ▶（塗りの三角形）を作成します。凹みを指定すると左辺がへこんだ形になります。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createTriangleMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.width;
        var height = placement.height;
        var inset = Math.max(0, Math.min(markSettings.insetPt, width * MAX_INSET_RATIO));
        var leftX = centerX - width / 2;
        var rightX = centerX + width / 2;
        var halfHeight = height / 2;

        var points = (inset > 0)
            ? [[rightX, centerY], [leftX, centerY + halfHeight], [leftX + inset, centerY], [leftX, centerY - halfHeight]]
            : [[rightX, centerY], [leftX, centerY + halfHeight], [leftX, centerY - halfHeight]];

        var trianglePath = createFilledPath(points);
        if (markSettings.roundCorners) {
            applyRoundCornersEffect(trianglePath, getTriangleCornerRadius(width, height, inset));
        }
        return trianglePath;
    }

    /**
     * 矢じりの奥行き（水平方向の長さ）を求めます。
     * 高さの半分を基本としつつ、矢じりが幅からはみ出さないよう上限を設けます。
     *
     * @param {number} width - 幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @returns {number} 矢じりの奥行き（pt）。
     */
    function getArrowHeadDepth(width, height) {
        var headDepth = height / 2;
        if (width > 0) headDepth = Math.min(headDepth, width * MAX_ARROW_HEAD_RATIO);
        return headDepth;
    }

    /**
     * →（線の矢印）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createArrowMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var headHalfHeight = placement.height / 2;
        var headDepth = getArrowHeadDepth(width, placement.height);
        var tipX = centerX + width / 2;

        var shaft = createStrokedLine([centerX - width / 2, centerY], [tipX, centerY], markSettings);

        var head = createStrokedPath([
            [tipX - headDepth, centerY + headHalfHeight],
            [tipX, centerY],
            [tipX - headDepth, centerY - headHalfHeight]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true });

        return groupPageItems([shaft, head]);
    }

    /**
     * ➡（塗りの矢印）を作成します。
     * 高さ（％）は使わず、線幅を軸の太さとし、矢じりの天地を線幅から求めます。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createSolidArrowMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var shaftWidthPt = markSettings.strokeWidthPt;
        var height = shaftWidthPt * SOLID_ARROW_HEIGHT_TO_STROKE_RATIO;
        var headHalfHeight = height / 2;
        var headDepth = getArrowHeadDepth(width, height);
        var tipX = centerX + width / 2;

        var shaft = createStrokedPath([
            [centerX - width / 2, centerY],
            [tipX - headDepth, centerY]
        ], shaftWidthPt);

        var head = createFilledPath([
            [tipX - headDepth, centerY - headHalfHeight],
            [tipX, centerY],
            [tipX - headDepth, centerY + headHalfHeight]
        ]);

        return mergeSolidArrowParts(shaft, head);
    }

    /**
     * ➡ の軸と矢じりを、アウトライン化と合体で1つのパスにまとめます。
     * 選択状態を使うメニューコマンドを呼ぶため、呼び出し側（createMarks）で選択を退避・復元します。
     * メニューコマンドは失敗しても例外の中身が読めないので、ここでは受け止めてフォールバックへ回します。
     * 合成できなかった場合も、軸と矢じりを取り残さないよう1つのグループにまとめて返します。
     *
     * @param {PathItem} shaft - 軸（線）。
     * @param {PathItem} head - 矢じり（塗り）。
     * @returns {PageItem} 合成後のアイテム。
     */
    function mergeSolidArrowParts(shaft, head) {
        var activeDoc = app.activeDocument;
        var outlinedShaft = shaft;

        try {
            activeDoc.selection = [shaft];
            app.executeMenuCommand("Live Outline Stroke");
            if (activeDoc.selection.length > 0) {
                outlinedShaft = activeDoc.selection[0];
            }

            activeDoc.selection = [outlinedShaft, head];
            app.executeMenuCommand("group");
            app.executeMenuCommand("Live Pathfinder Add");

            /* 1つにまとまったときだけ成功とみなす / Treat it as merged only when a single item is left */
            if (activeDoc.selection.length === 1) {
                return activeDoc.selection[0];
            }
        } catch (e) {
            logScriptError(getLabel("log.mergeSolidArrow"), e);
        }

        /* メニューコマンドが効かなかった場合のフォールバック / Fallback when the menu commands did not take effect */
        var fallbackGroup = null;
        runSafely(function () {
            fallbackGroup = groupPageItems([outlinedShaft, head]);
        }, getLabel("log.mergeSolidArrow"));

        return fallbackGroup || outlinedShaft;
    }

    /**
     * ＿\（横線＋斜線）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createArrowSlashMark(placement, markSettings) {
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var rise = placement.height / 2;
        var leftX = placement.centerX - width / 2;
        var tipX = placement.centerX + width / 2;

        var slashAngle = markSettings.angleDeg;
        if (isNaN(slashAngle) || slashAngle <= 0) slashAngle = SLASH_ANGLE_DEFAULT;
        if (slashAngle >= SLASH_ANGLE_MAX) slashAngle = SLASH_ANGLE_MAX;

        var slashDx = rise / Math.tan(slashAngle * Math.PI / 180);
        if (!isFinite(slashDx) || slashDx <= 0) slashDx = rise;

        var slashTopX = tipX - slashDx;
        if (slashTopX <= leftX) {
            slashTopX = leftX + Math.max(markSettings.strokeWidthPt, width * 0.15);
        }

        return createStrokedPath([
            [leftX, centerY],
            [tipX, centerY],
            [slashTopX, centerY + rise]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true, flatWhenNotRound: true });
    }

    /**
     * 天地を水平にした > の腕を1本作成します。
     *
     * @param {number} startX - 左端の X 座標。
     * @param {number} startY - 左端の Y 座標。
     * @param {number} tipX - 先端の X 座標。
     * @param {number} tipY - 先端の Y 座標。
     * @param {number} thickness - 腕の太さ（pt）。
     * @returns {PathItem} 作成したパス。
     */
    function createFlatChevronArm(startX, startY, tipX, tipY, thickness) {
        return createFilledPath([
            [startX, startY],
            [startX + thickness, startY],
            [tipX, tipY],
            [tipX - thickness, tipY]
        ]);
    }

    /**
     * > を1つ作成します。［天地を水平に］がONなら塗りの形状、OFFなら線で描きます。
     *
     * @param {number} centerX - 中心の X 座標。
     * @param {number} centerY - 中心の Y 座標。
     * @param {number} width - 山形1つ分の幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createSingleChevron(centerX, centerY, width, height, markSettings) {
        var leftX = centerX - width / 2;
        var tipX = centerX + width / 2;
        var halfHeight = height / 2;

        if (markSettings.flatChevron) {
            return groupPageItems([
                createFlatChevronArm(leftX, centerY + halfHeight, tipX, centerY, markSettings.strokeWidthPt),
                createFlatChevronArm(leftX, centerY - halfHeight, tipX, centerY, markSettings.strokeWidthPt)
            ]);
        }

        return createStrokedPath([
            [leftX, centerY + halfHeight],
            [tipX, centerY],
            [leftX, centerY - halfHeight]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true });
    }

    /**
     * >（山形）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createChevronMark(placement, markSettings) {
        return createSingleChevron(placement.centerX, placement.centerY, placement.width, placement.height, markSettings);
    }

    /**
     * >>（二重の山形）を作成します。幅は山形1つ分で、2つの間を［間隔］だけ空けます。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createDoubleChevronMark(placement, markSettings) {
        var centerOffset = (placement.width + markSettings.chevronGapPt) / 2;
        return groupPageItems([
            createSingleChevron(placement.centerX - centerOffset, placement.centerY, placement.width, placement.height, markSettings),
            createSingleChevron(placement.centerX + centerOffset, placement.centerY, placement.width, placement.height, markSettings)
        ]);
    }

    /**
     * ─（横線）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createDashMark(placement, markSettings) {
        var halfWidth = placement.width / 2;
        return createStrokedLine(
            [placement.centerX - halfWidth, placement.centerY],
            [placement.centerX + halfWidth, placement.centerY],
            markSettings
        );
    }

    /**
     * ＋（十字）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createPlusMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var halfWidth = placement.width / 2;

        return groupPageItems([
            createStrokedLine([centerX - halfWidth, centerY], [centerX + halfWidth, centerY], markSettings),
            createStrokedLine([centerX, centerY + halfWidth], [centerX, centerY - halfWidth], markSettings)
        ]);
    }

    /**
     * ×（斜め十字）を作成します。＋ を45°回転させた形です。
     *
     * ほかの形状と違い、幅は実寸ではなく ＋ と同じ腕の長さを表します（描画幅は幅 × 0.707）。
     * 実寸で揃えると ＋ より大きく見えるため、意図的にこの基準にしています。
     *
     * Unlike the other shapes, the width here is the arm length shared with the plus sign, not the drawn width.
     * Matching the drawn width would make it look larger than the plus sign, so this is intentional.
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createMultiplyMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var diagonalOffset = (placement.width / 2) * Math.SQRT2 / 2;

        return groupPageItems([
            createStrokedLine([centerX - diagonalOffset, centerY + diagonalOffset], [centerX + diagonalOffset, centerY - diagonalOffset], markSettings),
            createStrokedLine([centerX - diagonalOffset, centerY - diagonalOffset], [centerX + diagonalOffset, centerY + diagonalOffset], markSettings)
        ]);
    }

    // =========================================
    // 記号の作成 / Creating marks
    // =========================================

    /**
     * 隣り合う2つのオブジェクトの間に記号を1つ作成し、［左右逆］がONなら反転します。
     *
     * @param {PageItem} leftItem - 左側のアイテム。
     * @param {PageItem} rightItem - 右側のアイテム。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。アキがない場合は null。
     */
    function createMarkBetweenItems(leftItem, rightItem, markSettings) {
        var placement = computePlacementBetweenItems(leftItem, rightItem, markSettings.heightPercent, markSettings.widthPt, markSettings.offsetX, markSettings.offsetY);
        if (!placement) return null;

        var markItem = markSettings.shapeConfig.createMark(placement, markSettings);
        if (markSettings.mirror) {
            markItem.resize(-100, 100, true, true, true, true, 100, Transformation.CENTER);
        }
        return markItem;
    }

    /**
     * 隣り合うすべての組に記号を作成します。
     *
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem[]} 作成したアイテムの配列。
     */
    function createMarks(sortedItems, markSettings) {
        var createdItems = [];

        /* ➡ だけは選択状態を使うメニューコマンドを呼ぶため、ここで1回だけ退避・復元する / Only the solid arrow drives menu commands, so save and restore the selection once */
        var previousSelection = markSettings.shapeConfig.changesSelection
            ? collectionToArray(app.activeDocument.selection)
            : null;

        try {
            for (var i = 0; i < sortedItems.length - 1; i++) {
                var markItem = createMarkBetweenItems(sortedItems[i], sortedItems[i + 1], markSettings);
                if (markItem) createdItems.push(markItem);
            }
        } finally {
            if (previousSelection) restoreSelection(previousSelection);
        }
        return createdItems;
    }

    // =========================================
    // ダイアログの構築 / Building the dialog
    // =========================================

    /**
     * 項目名・入力欄・単位を1行にまとめて追加します。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {string} labelPath - 項目名のドットパス。
     * @param {string} unitText - 入力欄の右に置く単位表記。
     * @param {string} [initialText] - 入力欄の初期値。省略時は空欄。
     * @returns {EditText} 追加した入力欄。行のグループは rowGroup、∧∨は stepperGroup、増減の設定は stepOptions で参照できます
     *     （増減量・下限は setupFieldStepper() で設定します）。
     */
    function addLabeledField(parentPanel, labelPath, unitText, initialText) {
        var fieldRowGroup = parentPanel.add("group");
        setupRow(fieldRowGroup, "left", FIELD_ROW_SPACING);

        var fieldCaption = fieldRowGroup.add("statictext", undefined, labelText(labelPath));
        fieldCaption.preferredSize = [LABEL_COLUMN_WIDTH, -1];
        fieldCaption.justify = "right";

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var stepOptions = { step: 1 };
        var inputField;
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);
        inputField = stepperInputGroup.add("edittext", undefined, initialText || "");
        inputField.characters = FIELD_CHARACTERS;
        inputField.rowGroup = fieldRowGroup;
        inputField.stepperGroup = stepperGroup;
        inputField.stepOptions = stepOptions;
        bindSteppedArrowKeys(inputField, stepperGroup);

        fieldRowGroup.add("statictext", undefined, unitText);

        return inputField;
    }

    /**
     * ［形状］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildShapePanel(parentColumn, dialogControls) {
        var shapePanel = parentColumn.add("panel", undefined, getLabel("panel.shape"));
        setupPanel(shapePanel, RADIO_SPACING);

        dialogControls.shapeRadios = {};
        for (var i = 0; i < SHAPE_ORDER.length; i++) {
            var shapeConfig = SHAPE_CONFIG[SHAPE_ORDER[i]];
            var shapeRadio = shapePanel.add("radiobutton", undefined, shapeConfig.symbol);
            if (shapeConfig.helpTip) shapeRadio.helpTip = shapeConfig.helpTip;
            dialogControls.shapeRadios[SHAPE_ORDER[i]] = shapeRadio;
        }
        dialogControls.shapeRadios[SHAPE_ORDER[0]].value = true;

        var mirrorGroup = shapePanel.add("group");
        setupRow(mirrorGroup, "left");
        mirrorGroup.margins = [0, MIRROR_ROW_TOP_MARGIN, 0, 0];

        dialogControls.mirrorCheckbox = mirrorGroup.add("checkbox", undefined, getLabel("checkbox.mirrorHorizontal"));
        dialogControls.mirrorCheckbox.helpTip = getLabel("tooltip.mirrorHorizontal");
    }

    /**
     * ［先端］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildCapStylePanel(parentColumn, dialogControls) {
        var capStylePanel = parentColumn.add("panel", undefined, getLabel("panel.capStyle"));
        setupPanel(capStylePanel, RADIO_SPACING);

        dialogControls.capStylePanel = capStylePanel;
        dialogControls.capNoneRadio = capStylePanel.add("radiobutton", undefined, getLabel("radio.capNone"));
        dialogControls.capNoneRadio.helpTip = getLabel("tooltip.capNone");
        dialogControls.capRoundRadio = capStylePanel.add("radiobutton", undefined, getLabel("radio.capRound"));
        dialogControls.capRoundRadio.helpTip = getLabel("tooltip.capRound");
        dialogControls.capNoneRadio.value = true;
    }

    /**
     * ［オプション］パネルを作成します。
     * 入力欄の初期値は、形状に合わせて applyShapeDefaults() で入れます。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildOptionsPanel(parentColumn, dialogControls) {
        var optionsPanel = parentColumn.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, FIELD_ROW_SPACING);

        dialogControls.heightField = addLabeledField(optionsPanel, "fieldLabel.height", "%");
        dialogControls.heightField.helpTip = getLabel("tooltip.height", [MAX_HEIGHT_PERCENT]);
        dialogControls.heightField.active = true;

        dialogControls.widthField = addLabeledField(optionsPanel, "fieldLabel.width", rulerUnitInfo.label);
        dialogControls.widthField.helpTip = getLabel("tooltip.width");

        dialogControls.insetField = addLabeledField(optionsPanel, "fieldLabel.inset", rulerUnitInfo.label);
        dialogControls.insetField.helpTip = getLabel("tooltip.inset", [roundDisplayValue(MAX_INSET_RATIO * 100)]);

        dialogControls.gapField = addLabeledField(optionsPanel, "fieldLabel.gap", rulerUnitInfo.label);
        dialogControls.gapField.helpTip = getLabel("tooltip.gap");

        dialogControls.strokeField = addLabeledField(optionsPanel, "fieldLabel.strokeWidth", strokeUnitInfo.label);

        dialogControls.angleField = addLabeledField(optionsPanel, "fieldLabel.angle", "°");
        dialogControls.angleField.helpTip = getLabel("tooltip.angle", [SLASH_ANGLE_MAX]);

        var optionCheckboxGroup = optionsPanel.add("group");
        optionCheckboxGroup.orientation = "column";
        optionCheckboxGroup.alignChildren = ["left", "top"];
        optionCheckboxGroup.spacing = FIELD_ROW_SPACING;
        optionCheckboxGroup.margins = [OPTION_CHECKBOX_INDENT, OPTION_CHECKBOX_TOP_MARGIN, 0, 0];

        dialogControls.flatChevronCheckbox = optionCheckboxGroup.add("checkbox", undefined, getLabel("checkbox.flatChevron"));
        dialogControls.flatChevronCheckbox.helpTip = getLabel("tooltip.flatChevron");

        dialogControls.roundCornersCheckbox = optionCheckboxGroup.add("checkbox", undefined, getLabel("checkbox.roundCorners"));
        dialogControls.roundCornersCheckbox.helpTip = getLabel("tooltip.roundCorners");
        dialogControls.roundCornersCheckbox.value = true;
    }

    /**
     * ［位置調整］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildPositionPanel(parentColumn, dialogControls) {
        var positionPanel = parentColumn.add("panel", undefined, getLabel("panel.position"));
        setupPanel(positionPanel, FIELD_ROW_SPACING);

        dialogControls.offsetXField = addLabeledField(positionPanel, "fieldLabel.offsetX", rulerUnitInfo.label, "0");
        dialogControls.offsetXField.helpTip = getLabel("tooltip.offsetX");
        dialogControls.offsetYField = addLabeledField(positionPanel, "fieldLabel.offsetY", rulerUnitInfo.label, "0");
        dialogControls.offsetYField.helpTip = getLabel("tooltip.offsetY");
    }

    /**
     * ダイアログの中身を組み立て、コントロールをまとめて返します。
     *
     * @param {Window} markDialog - 対象のダイアログ。
     * @returns {object} すべてのコントロールを持つオブジェクト。
     */
    function buildDialogControls(markDialog) {
        var columnsGroup = markDialog.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumn = columnsGroup.add("group");
        setupColumn(leftColumn, COLUMN_PANEL_SPACING);

        var rightColumn = columnsGroup.add("group");
        setupColumn(rightColumn, COLUMN_PANEL_SPACING);

        var dialogControls = {};
        buildShapePanel(leftColumn, dialogControls);
        buildCapStylePanel(leftColumn, dialogControls);
        buildOptionsPanel(rightColumn, dialogControls);
        buildPositionPanel(rightColumn, dialogControls);

        /* ［プレビュー］とボタンの行。［キャンセル］は name: "cancel" の既定動作で閉じ、プレビューは onClose で片付ける
           Preview and button row. Cancel closes via name: "cancel"; the preview is cleaned up in onClose */
        var buttonRow = addButtonRow(markDialog);
        dialogControls.previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        centerButtonRowIfRightOnly(buttonRow);
        dialogControls.btnOK = btnOK;
        return dialogControls;
    }

    // =========================================
    // ダイアログの状態と入力値 / Dialog state and input values
    // =========================================

    /**
     * 選択中の形状キーを返します。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @returns {string} SHAPE_CONFIG のキー。
     */
    function getSelectedShapeKey(dialogControls) {
        for (var i = 0; i < SHAPE_ORDER.length; i++) {
            if (dialogControls.shapeRadios[SHAPE_ORDER[i]].value) return SHAPE_ORDER[i];
        }
        return SHAPE_ORDER[0];
    }

    /**
     * ［先端］を［丸型］か［なし］に切り替えます。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {boolean} isRound - ［丸型］にするなら true。
     * @returns {void}
     */
    function setRoundCaps(dialogControls, isRound) {
        dialogControls.capRoundRadio.value = isRound;
        dialogControls.capNoneRadio.value = !isRound;
    }

    /**
     * チェックボックスの有効・無効を切り替え、無効にしたときはOFFに戻します。
     *
     * @param {Checkbox} checkbox - 対象のチェックボックス。
     * @param {boolean} isAvailable - 有効にするなら true。
     * @returns {void}
     */
    function setCheckboxAvailability(checkbox, isAvailable) {
        checkbox.enabled = isAvailable;
        if (!isAvailable) checkbox.value = false;
    }

    /**
     * 入力欄の行（項目名・∧∨・入力欄・単位）をまとめて有効／無効にし、∧∨を描き直します。
     *
     * @param {EditText} inputField - addLabeledField() で作った入力欄。
     * @param {boolean} isEnabled - 有効にするなら true。
     * @returns {void}
     */
    function setFieldRowEnabled(inputField, isEnabled) {
        inputField.rowGroup.enabled = isEnabled;
        redrawSteppersIn(inputField.rowGroup);
    }

    /**
     * 形状に応じて、各コントロールの有効・無効を切り替えます。
     * 入力欄は、項目名と単位表記もまとめてグレーにするため行ごと切り替えます。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {object} shapeConfig - 形状の設定。
     * @returns {void}
     */
    function applyShapeEnabledStates(dialogControls, shapeConfig) {
        dialogControls.capStylePanel.enabled = shapeConfig.enableCapPanel;
        if (!shapeConfig.enableCapPanel) setRoundCaps(dialogControls, false);

        setFieldRowEnabled(dialogControls.heightField, shapeConfig.enableHeightInput);
        setFieldRowEnabled(dialogControls.strokeField, shapeConfig.enableStrokeInput);
        setFieldRowEnabled(dialogControls.gapField, shapeConfig.enableGap);
        setFieldRowEnabled(dialogControls.insetField, shapeConfig.enableInset);
        setFieldRowEnabled(dialogControls.angleField, shapeConfig.enableAngle);

        setCheckboxAvailability(dialogControls.flatChevronCheckbox, shapeConfig.enableFlatChevron);
        setCheckboxAvailability(dialogControls.mirrorCheckbox, shapeConfig.enableMirror);
        setCheckboxAvailability(dialogControls.roundCornersCheckbox, shapeConfig.enableRoundCorners);
    }

    /**
     * 入力エラーを知らせ、対象の入力欄にフォーカスを移します。
     *
     * @param {EditText} invalidField - 対象の入力欄。
     * @param {string} alertPath - 表示するメッセージのドットパス。
     * @param {boolean} showAlert - 警告を表示するなら true。
     * @param {Array} [replacements] - メッセージに差し込む値。
     * @returns {null} 呼び出し元がそのまま返せるよう null を返します。
     */
    function reportInvalidValue(invalidField, alertPath, showAlert, replacements) {
        if (showAlert) {
            alert(getLabel(alertPath, replacements));
            invalidField.active = true;
        }
        return null;
    }

    /**
     * ダイアログの入力値をまとめて読み取り、検証します。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {boolean} showAlert - 不正な値のときに警告を表示するなら true。
     * @returns {object} 検証済みの入力値（寸法は pt）。不正な場合は null。
     */
    function readMarkSettings(dialogControls, showAlert) {
        var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];
        var heightPercent = null;
        var widthPt = parseFieldToPt(dialogControls.widthField, rulerUnitInfo, false);
        var insetPt = 0;
        var offsetX = parseFieldToPt(dialogControls.offsetXField, rulerUnitInfo, true);
        var offsetY = parseFieldToPt(dialogControls.offsetYField, rulerUnitInfo, true);
        var strokeWidthPt = parseFieldToPt(dialogControls.strokeField, strokeUnitInfo, false);
        var angleDeg = shapeConfig.enableAngle ? parseFloat(dialogControls.angleField.text) : 0;
        var chevronGapPt = parseFieldToPt(dialogControls.gapField, rulerUnitInfo, true);

        if (shapeConfig.enableInset) {
            insetPt = parseFieldToPt(dialogControls.insetField, rulerUnitInfo, false);
            if (isNaN(insetPt)) {
                return reportInvalidValue(dialogControls.insetField, "alert.invalidValue", showAlert);
            }
            /* 幅が明示されている場合は、その割合まで凹みを抑える / Clamp the inset when the width is set explicitly */
            if (widthPt > 0) {
                insetPt = Math.min(insetPt, widthPt * MAX_INSET_RATIO);
            }
        }

        if (shapeConfig.enableHeightInput) {
            heightPercent = parseFloat(dialogControls.heightField.text);
            if (isNaN(heightPercent) || heightPercent <= 0) {
                return reportInvalidValue(dialogControls.heightField, "alert.positiveNumber", showAlert);
            }
            if (heightPercent > MAX_HEIGHT_PERCENT) {
                return reportInvalidValue(dialogControls.heightField, "alert.maxHeight", showAlert, [MAX_HEIGHT_PERCENT]);
            }
        }

        if (isNaN(widthPt)) {
            return reportInvalidValue(dialogControls.widthField, "alert.widthPositive", showAlert);
        }
        if (isNaN(offsetX)) {
            return reportInvalidValue(dialogControls.offsetXField, "alert.invalidValue", showAlert);
        }
        if (isNaN(offsetY)) {
            return reportInvalidValue(dialogControls.offsetYField, "alert.invalidValue", showAlert);
        }
        if (isNaN(strokeWidthPt)) {
            return reportInvalidValue(dialogControls.strokeField, "alert.invalidValue", showAlert);
        }
        /* 塗りだけの形状は線幅を使わないので、下限を求めない / Fill-only shapes ignore the stroke width, so the minimum does not apply */
        if (shapeConfig.enableStrokeInput && strokeWidthPt < MIN_STROKE_WIDTH_PT) {
            var minStrokeText = roundDisplayValueForUnit(convertPtToUnitValue(MIN_STROKE_WIDTH_PT, strokeUnitInfo), strokeUnitInfo);
            return reportInvalidValue(dialogControls.strokeField, "alert.strokePositive", showAlert, [minStrokeText, strokeUnitInfo.label]);
        }
        if (shapeConfig.enableAngle && isNaN(angleDeg)) {
            return reportInvalidValue(dialogControls.angleField, "alert.invalidValue", showAlert);
        }

        return {
            shapeConfig: shapeConfig,
            heightPercent: heightPercent,
            widthPt: widthPt,
            insetPt: insetPt,
            offsetX: offsetX,
            offsetY: offsetY,
            strokeWidthPt: strokeWidthPt,
            angleDeg: angleDeg,
            chevronGapPt: isNaN(chevronGapPt) ? 0 : chevronGapPt,
            roundEnds: dialogControls.capRoundRadio.value,
            roundCorners: dialogControls.roundCornersCheckbox.value,
            flatChevron: dialogControls.flatChevronCheckbox.value,
            mirror: dialogControls.mirrorCheckbox.value
        };
    }

    /**
     * 入力欄の∧∨（と↑↓キー）の増減量・下限を、単位に合わせて設定します。
     * 基本の増減量は単位に合わせて決まり（inch などは細かい刻み）、Shift は10の倍数へ、Option（Alt）は0.1ずつ増減します。
     * 増減したあとは単位の表示桁に丸め、入力欄の onChanging を呼んで手入力と同じ後処理を通します。
     *
     * @param {EditText} editText - addLabeledField() で作った入力欄。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
     * @param {number} [maxValue] - 上限。省略時は上限なし。
     * @returns {void}
     */
    function setupFieldStepper(editText, allowNegative, unitInfo, maxValue) {
        var stepOptions = editText.stepOptions;
        stepOptions.step = getArrowKeyStep(unitInfo);
        if (!allowNegative) stepOptions.min = 0;
        if (typeof maxValue === "number") stepOptions.max = maxValue;
        stepOptions.onStep = function (steppedField) {
            var value = parseFloat(steppedField.text);
            if (!isNaN(value)) steppedField.text = String(roundDisplayValueForUnit(value, unitInfo));
            if (steppedField.onChanging) steppedField.onChanging();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認してダイアログを表示します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        var activeDoc = app.activeDocument;

        /* ロック・非表示のレイヤーには作成できないので、先に知らせる / Nothing can be created on a locked or hidden layer */
        if (activeDoc.activeLayer.locked || !activeDoc.activeLayer.visible) {
            alert(getLabel("alert.lockedLayer"));
            return;
        }

        var sortedItems = collectionToArray(activeDoc.selection);
        if (sortedItems.length < 2) {
            alert(getLabel("alert.selectTwoObjects"));
            return;
        }
        sortItemsLeftToRight(sortedItems);

        /* どの組にもアキがなければ記号を作成できないので、ダイアログを出す前に知らせる / Fail early when no pair has a usable gap */
        var narrowestGap = measureNarrowestGap(sortedItems);
        if (!narrowestGap) {
            alert(getLabel("alert.noGap"));
            return;
        }

        showMarkDialog(activeDoc, sortedItems, narrowestGap);
    }

    /**
     * ダイアログを表示し、プレビューと［OK］で記号を作成します。
     *
     * @param {Document} activeDoc - 対象のドキュメント。
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @param {object} narrowestGap - measureNarrowestGap() の戻り値（幅の自動計算に使います）。
     * @returns {void}
     */
    function showMarkDialog(activeDoc, sortedItems, narrowestGap) {
        var markDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(markDialog);

        var dialogControls = buildDialogControls(markDialog);

        /* 幅を手入力したかどうか（手入力後は自動計算しない）/ Whether the width was typed in (auto-calculation stops once it is) */
        var widthManuallySet = false;

        /* プレビューで作成したアイテム / Items created for the preview */
        var previewItems = [];

        /**
         * 現在の形状と高さから、幅の自動計算値（pt）を求めます。
         *
         * @returns {number} 幅（pt）。高さが不正で求められない場合は null。
         */
        function computeAutoWidthPt() {
            var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];
            var heightPercent = parseFloat(dialogControls.heightField.text);
            if (shapeConfig.enableHeightInput && (isNaN(heightPercent) || heightPercent <= 0)) return null;

            var basePt = (shapeConfig.autoWidth.basis === "gap")
                ? narrowestGap.minGapWidth
                : narrowestGap.minTotalHeight * (heightPercent / 100);
            return roundDisplayValue(basePt * shapeConfig.autoWidth.ratio);
        }

        /**
         * 幅の入力欄に、自動計算した値を表示します。
         *
         * @returns {void}
         */
        function applyAutoWidthToField() {
            var autoWidthPt = computeAutoWidthPt();
            if (autoWidthPt === null) return;
            setFieldFromPt(dialogControls.widthField, autoWidthPt, rulerUnitInfo);
        }

        /**
         * 幅を自動計算に戻します。
         *
         * @returns {void}
         */
        function resetWidthToAuto() {
            widthManuallySet = false;
            applyAutoWidthToField();
        }

        /**
         * 実際に使われる幅（pt）を求めます。手入力があればその値を優先します。
         *
         * @returns {number} 幅（pt）。求められない場合は 0。
         */
        function getEffectiveWidthPt() {
            var typedWidthPt = parseFieldToPt(dialogControls.widthField, rulerUnitInfo, false);
            if (typedWidthPt > 0) return typedWidthPt;

            var autoWidthPt = computeAutoWidthPt();
            return (autoWidthPt > 0) ? autoWidthPt : 0;
        }

        /**
         * 凹みが幅の割合を超えないよう、入力欄の値を抑えます。
         *
         * @returns {void}
         */
        function clampInsetToWidth() {
            var insetValue = parseFloat(dialogControls.insetField.text);
            if (isNaN(insetValue) || insetValue < 0) return;

            var effectiveWidthPt = getEffectiveWidthPt();
            var maxInsetPt = effectiveWidthPt * MAX_INSET_RATIO;
            if (effectiveWidthPt > 0 && convertValueToPt(insetValue, rulerUnitInfo) > maxInsetPt) {
                setFieldFromPt(dialogControls.insetField, maxInsetPt, rulerUnitInfo);
            }
        }

        /**
         * 選択中の形状に合わせて、入力欄の初期値とコントロールの有効・無効を更新します。
         * 形状ごとに適切な値が違うため、切り替えるたびにその形状の初期値へ戻します。
         *
         * @returns {void}
         */
        function applyShapeDefaults() {
            var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];

            if (shapeConfig.enableHeightInput) {
                dialogControls.heightField.text = String(shapeConfig.defaultHeightPercent);
            }
            resetWidthToAuto();
            setFieldFromPt(dialogControls.gapField, shapeConfig.defaultGap, rulerUnitInfo);
            setFieldFromPt(dialogControls.insetField, 0, rulerUnitInfo);
            setFieldFromPt(dialogControls.strokeField, shapeConfig.defaultStrokePt, strokeUnitInfo);
            dialogControls.angleField.text = String(shapeConfig.defaultAngle);

            applyShapeEnabledStates(dialogControls, shapeConfig);
        }

        /**
         * 形状を切り替えたときに、初期値と位置調整をリセットします。
         *
         * @returns {void}
         */
        function handleShapeChange() {
            dialogControls.offsetXField.text = "0";
            dialogControls.offsetYField.text = "0";
            applyShapeDefaults();
            updatePreview();
        }

        /**
         * プレビューで作成したアイテムを削除します。
         *
         * @returns {void}
         */
        function removePreview() {
            if (previewItems.length === 0) return;

            for (var i = 0; i < previewItems.length; i++) {
                removeItemSafely(previewItems[i], getLabel("log.removePreviewItem"));
            }
            previewItems = [];
            app.redraw();
        }

        /**
         * プレビューを作り直します。
         *
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            if (!dialogControls.previewCheckbox.value) return;

            var markSettings = readMarkSettings(dialogControls, false);
            if (!markSettings) return;

            previewItems = createMarks(sortedItems, markSettings);
            app.redraw();
        }

        /**
         * ［OK］で記号を確定します。プレビューがあれば、それをそのまま結果にします。
         *
         * @returns {void}
         */
        function confirmMarks() {
            var markSettings = readMarkSettings(dialogControls, true);
            if (!markSettings) return;

            /* onClose で消されないよう、プレビューを手放してから閉じる / Take over the preview so onClose does not remove it */
            var createdItems = previewItems;
            previewItems = [];
            markDialog.close(1);

            if (createdItems.length === 0) {
                createdItems = createMarks(sortedItems, markSettings);
                if (createdItems.length === 0) {
                    alert(getLabel("alert.noGap"));
                    return;
                }
            }
            activeDoc.selection = createdItems;
        }

        /**
         * 入力欄の変更と∧∨・矢印キーの増減を登録します。
         *
         * @returns {void}
         */
        function bindFieldEvents() {
            dialogControls.heightField.onChanging = function () {
                if (!widthManuallySet) applyAutoWidthToField();
                updatePreview();
            };

            dialogControls.widthField.onChanging = function () {
                if (dialogControls.widthField.text === "") {
                    resetWidthToAuto();
                } else {
                    widthManuallySet = true;
                }
                updatePreview();
            };

            dialogControls.insetField.onChanging = function () {
                clampInsetToWidth();
                updatePreview();
            };

            var previewOnlyFields = [
                dialogControls.gapField,
                dialogControls.strokeField,
                dialogControls.angleField,
                dialogControls.offsetXField,
                dialogControls.offsetYField
            ];
            for (var i = 0; i < previewOnlyFields.length; i++) {
                previewOnlyFields[i].onChanging = updatePreview;
            }

            /* ％と度は単位を持たないため、増減量は1固定 / Percent and degrees have no unit, so they step by one */
            setupFieldStepper(dialogControls.heightField, false, null, MAX_HEIGHT_PERCENT);
            setupFieldStepper(dialogControls.angleField, false, null, SLASH_ANGLE_MAX);
            setupFieldStepper(dialogControls.widthField, false, rulerUnitInfo);
            setupFieldStepper(dialogControls.insetField, false, rulerUnitInfo);
            setupFieldStepper(dialogControls.strokeField, false, strokeUnitInfo);
            setupFieldStepper(dialogControls.gapField, true, rulerUnitInfo);
            setupFieldStepper(dialogControls.offsetXField, true, rulerUnitInfo);
            setupFieldStepper(dialogControls.offsetYField, true, rulerUnitInfo);
        }

        /**
         * 形状・先端・チェックボックスのクリックを登録します。
         *
         * @returns {void}
         */
        function bindOptionEvents() {
            for (var i = 0; i < SHAPE_ORDER.length; i++) {
                dialogControls.shapeRadios[SHAPE_ORDER[i]].onClick = handleShapeChange;
            }

            var previewToggles = [
                dialogControls.capNoneRadio,
                dialogControls.capRoundRadio,
                dialogControls.mirrorCheckbox,
                dialogControls.flatChevronCheckbox,
                dialogControls.roundCornersCheckbox,
                dialogControls.previewCheckbox
            ];
            for (var j = 0; j < previewToggles.length; j++) {
                previewToggles[j].onClick = updatePreview;
            }
        }

        /**
         * キーボードショートカット（F：先端なし、R：丸型、V：左右逆）を登録します。
         * 無効なときは false を返し、文字をそのまま通します。
         *
         * @returns {void}
         */
        function bindShortcutKeys() {
            addKeyShortcuts(markDialog, {
                F: function () {
                    if (!dialogControls.capStylePanel.enabled) return false;
                    setRoundCaps(dialogControls, false);
                    updatePreview();
                    return true;
                },
                R: function () {
                    if (!dialogControls.capStylePanel.enabled) return false;
                    setRoundCaps(dialogControls, true);
                    updatePreview();
                    return true;
                },
                /* 反転して onClick（updatePreview）を呼ぶ / Toggle, then its onClick refreshes the preview */
                V: function () {
                    if (!dialogControls.mirrorCheckbox.enabled) return false;
                    return dialogControls.mirrorCheckbox;
                }
            }, {
                numericFields: [
                    dialogControls.heightField, dialogControls.widthField, dialogControls.insetField,
                    dialogControls.gapField, dialogControls.strokeField, dialogControls.angleField,
                    dialogControls.offsetXField, dialogControls.offsetYField
                ]
            });
        }

        bindFieldEvents();
        bindOptionEvents();
        bindShortcutKeys();
        dialogControls.btnOK.onClick = confirmMarks;
        markDialog.onClose = removePreview;

        applyShapeDefaults();

        /* 表示前にレイアウトを確定して描画欠けを防ぐ / Settle the layout before showing to avoid partial rendering */
        markDialog.layout.layout(true);
        markDialog.layout.resize();

        prepareDialogWindow(markDialog, SCRIPT_NAME);
        markDialog.show();
    }

    main();

})();
