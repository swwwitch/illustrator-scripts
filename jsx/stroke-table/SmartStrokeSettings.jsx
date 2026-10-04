#target illustrator
#targetengine "SmartStrokeSettingsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスに、よく使う矢印と線の設定（線幅・線端・角の形状・カラー・破線）をまとめて適用します。
矢印はDOMから操作できないため、一時アクションを生成して実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartStrokeSettings.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1726fc0f8dc9

### Overview

Applies a favorite arrowhead together with the stroke settings — weight, cap, corner, color and dashes — to the selected paths.
Arrowheads cannot be reached from the DOM, so a temporary action is generated and played instead.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartStrokeSettings.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartStrokeSettings";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartStrokeSettings.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartStrokeSettings.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1726fc0f8dc9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 既定値 / Defaults */
    var DEFAULT_STROKE_WIDTH = 5;            /* 線幅の初期値（pt）/ initial stroke width (pt) */
    /* 一般の単位ごとの線幅の初期値（pt）。ここに無い単位は DEFAULT_STROKE_WIDTH（1px＝1pt）
       Initial stroke width (pt) per general unit; other units use DEFAULT_STROKE_WIDTH (1 px = 1 pt) */
    var DEFAULT_STROKE_WIDTH_BY_UNIT = { "mm": 0.25, "px": 1 };
    var DEFAULT_ARROW_NUMBER = 11;           /* 最初に選んでおく矢印の番号（FAVORITE_ARROWS から）/ arrowhead number selected at start (from FAVORITE_ARROWS) */
    var DEFAULT_ARROW_SCALE  = 100;          /* 倍率の初期値（ポップアップメニューの矢印にも使う）/ initial arrowhead scale, also used for the pop-up arrowheads */
    /* 線幅のポップアップメニューに並べる値（pt、Illustrator の線パネルと同じ）/ weights in the pop-up menu (pt, as in Illustrator's Stroke panel) */
    var STROKE_WIDTH_PRESETS = [0.25, 0.5, 0.75, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100];
    var DEFAULT_STROKE_CAP   = "butt";       /* 線端の初期値（butt / round / projecting）/ default cap */
    var DEFAULT_CORNER_JOIN  = "miter";      /* 角の形状の初期値（miter / round / bevel）/ default join */

    /* アイコンで出すよく使う矢印と倍率・先端位置。ここに無い矢印はポップアップメニューに並ぶ
       number 0 は「[なし]」（矢印を外す）。tipAlign は TIP_ALIGN_OPTIONS の key（atEnd / beyondEnd）、省略すると先端位置を変えない。
       strokeWidthScale を付けた矢印は、選んだときに線幅をその倍数にし、ほかの矢印を選ぶと元に戻す。
       strokeCap / cornerJoin（STROKE_CAP_OPTIONS・CORNER_JOIN_OPTIONS の key）を付けた矢印は、選んだときに線端・角の形状もそれにする。
       付けていない矢印を選ぶと、線端・角の形状は DEFAULT_STROKE_CAP・DEFAULT_CORNER_JOIN（線端なし・マイター）に戻る
       Favorite arrowheads shown as icons, with their scales and tip alignment. The rest go in the pop-up menu.
       Number 0 is [None], which removes the arrowheads. tipAlign is a TIP_ALIGN_OPTIONS key; omit it to leave the tip alone.
       An arrowhead with strokeWidthScale multiplies the weight while it is picked; picking another restores it.
       strokeCap / cornerJoin (keys of STROKE_CAP_OPTIONS / CORNER_JOIN_OPTIONS) also set the cap and corner when it is picked;
       picking an arrowhead without them returns the cap and corner to DEFAULT_STROKE_CAP / DEFAULT_CORNER_JOIN */
    var FAVORITE_ARROWS = [
        { number: 0,  scale: 100 },
        { number: 8,  scale: 25,  tipAlign: "atEnd", strokeWidthScale: 3 },  /* 線幅は 300% / weight at 300% */
        { number: 11, scale: 100, tipAlign: "atEnd" },
        { number: 13, scale: 100, strokeCap: "round", cornerJoin: "round" }, /* 丸型線端・ラウンド結合にする / round cap and join */
        { number: 21, scale: 33,  strokeCap: "round", cornerJoin: "round" }, /* 丸型線端・ラウンド結合にする / round cap and join */
        { number: 27, scale: 100, tipAlign: "beyondEnd" }
    ];

    /* 破線のパターン（線幅に対する倍率）。オンにしたときの分割数・間隔・線分の初期値に使う
       丸型線端では、見かけの線分は dash＋線幅、すき間は gap−線幅になる。ドット点線は丸型線端に固定
       Dash patterns (multiples of the stroke weight), used for the initial segments, gap and dash.
       With round caps a visible dash is dash + weight and a visible gap is gap - weight. Dotted always uses round caps */
    var DASH_PATTERNS = {
        dashed: { dash: 2, gap: 3 },         /* 破線 / dashed */
        dotted: { dash: 0, gap: 2 }          /* ドット点線 / dotted */
    };
    var DOT_GAP_PRECISION = 1000;            /* 両端を調整したドット点線の間隔の切り捨て単位（1/1000 pt）/ truncation unit for adjusted dotted gaps */

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    var ACTION_SET_NAME = "SwwwitchTempStrokeSet";
    var ACTION_NAME     = "SwwwitchTempStroke";
    var ARROW_COUNT     = 39;                /* Illustrator の矢印の種類数 / number of Illustrator arrowheads */
    var UNIT_POINT      = 592476268;         /* ポイント / point（parameter /unit） */

    /* パラメータキー / Parameter keys（記録した .aia から採取 / taken from a recorded .aia） */
    var KEY_STROKE_WIDTH  = 2003072104;      /* 線幅 / stroke width */
    var KEY_CAP           = 1667330094;      /* 線端 / cap */
    var KEY_JOIN          = 1785686382;      /* 角の形状 / join */
    var KEY_DASH_INT      = 1684825454;      /* 破線（整数）/ dash (integer) */
    var KEY_DASH_BOOL     = 1684104298;      /* 破線（真偽）/ dash (boolean) */
    var KEY_ARROW_HEAD_1  = 1634231345;      /* ahd1: 始点の形状 / start arrowhead */
    var KEY_ARROW_HEAD_2  = 1634231346;      /* ahd2: 終点の形状 / end arrowhead */
    var KEY_ARROW_SCALE_1 = 1634951985;      /* asc1: 始点の倍率（推定値）/ start scale (estimated) */
    var KEY_ARROW_SCALE_2 = 1634951986;      /* asc2: 終点の倍率 / end scale */
    var KEY_TIP_ALIGN     = 1634230636;      /* ahal: 矢印の配置 / tip alignment */
    var KEY_STROKE_ALIGN  = 1634494318;      /* algn: 線の位置 / stroke alignment */

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

    /* コントロールの寸法 / Control metrics */
    var ROW_SPACING        = 8;              /* 行内の要素間隔 / spacing inside a row */
    var STEPPER_LABEL_GAP  = 0;              /* 項目名（コロン）と∧∨の間 / gap between a label's colon and the stepper */
    var LABEL_GAP          = 4;              /* 項目名（コロン）と右の入力欄・アイコンの間 / gap between a label's colon and its control */
    var ARROW_LABEL_WIDTH  = 40;             /* 矢印パネルの項目名の幅 / arrowhead panel label width */
    var FIELD_CHARACTERS   = 4;              /* 数値入力欄の文字数 / numeric field width */
    var PRESET_ICON_SIZE   = [20, 18];       /* プリセットの保存・削除アイコンの大きさ（部品の既定 24×22 よりひとまわり小さく）/ preset save and delete icon size */
    var SAVE_ICON_OPACITY  = 0.75;           /* 保存アイコンの濃さ（ほかのアイコンに対する比率）/ save icon opacity relative to the others */
    var CHOICE_ICON_SIZE   = [36, 26];       /* 選択肢アイコン（先端位置・両端を調整）の大きさ。4つともそろえる / shared size of the tip and adjust-ends icons */
    var ADJUST_ICON_SIZE   = CHOICE_ICON_SIZE; /* 両端を調整のアイコンの大きさ / adjust-ends icon size */
    var STROKE_ICON_SIZE   = [28, 22];       /* 線端・角の形状のアイコンの大きさ（ほかの選択肢アイコンよりひとまわり小さい）/ cap and corner icon size, a size smaller */
    var STROKE_ICON_INSET  = 3;              /* 線端・角の形状のアイコンの枠と図形の間（px）/ gap between the cap / corner icon frame and its shape */
    var TIP_ICON_SIZE      = CHOICE_ICON_SIZE; /* 先端位置のアイコンの大きさ / tip alignment icon size */
    var ARROW_ICON_COLUMNS = 2;              /* よく使う矢印のアイコンを横に並べる数（［なし］も含む）/ favorite arrowhead icons per row, [None] included */
    var ARROW_ICON_SIZE    = [76, 26];       /* よく使う矢印のアイコンの大きさ / favorite arrowhead icon size */
    var ARROW_ICON_SCALE   = 22 / 383;       /* よく使う矢印の絵の倍率（高さを詰めても絵の大きさは変えない）/ fixed drawing scale, so a shorter icon keeps the same arrow */
    var ARROW_ICONS_BOTTOM_MARGIN = 10;      /* よく使う矢印のアイコン（最後の矢印）の下の余白 / space below the favorite arrowhead icons */
    var ARROW_ICON_INSET   = 5;              /* よく使う矢印のアイコンの枠と図形の間（px）/ gap between the arrowhead icon frame and its shape */
    var ARROW_OPTIONS_TOP_MARGIN = 10;       /* 矢印のオプション（アイコンの行）の上の余白 / space above the arrowhead option icons */
    var LINK_CHAIN_RATIO   = 0.8;            /* リンクアイコンの鎖の大きさ（枠に対する比率）/ chain size relative to the link icon */
    var OPTION_TOGGLE_SIZE = [30, 30];       /* 終点も同じ・入れ替えのアイコンの大きさ / Same at end and Swap icon size */
    var ADJUST_ICON_INSET  = 3;              /* 両端を調整のアイコンの枠と図形の間（px）/ gap between the adjust-ends icon frame and its shapes */
    var COLOR_SWATCH_SIZE  = [40, 20];       /* 線の色見本の大きさ / stroke color swatch size */
    var WIDTH_POPUP_BUTTON_WIDTH  = 20;      /* 線幅の▼ボタンの幅 / width of the weight ▼ button */
    var WIDTH_POPUP_BUTTON_HEIGHT = 22;      /* 線幅の▼ボタンの高さ（入力欄の高さが分からないとき）/ its height when the field's is unknown */
    var WIDTH_POPUP_LIST_SIZE     = [90, 300]; /* 線幅のリストの大きさ / size of the weight list */
    var UNIT_FIELD_CHARACTERS = 5;           /* 単位を欄の中に入れる数値欄の文字数 / width of a field holding its unit */
    var SEPARATOR_BOTTOM_MARGIN = 5;         /* 区切り線の下の余白 / space below a separator */
    var SUB_PANEL_TOP_MARGIN = 10;           /* 入れ子のパネル（計算方法）の上の余白 / space above nested panels */
    var PRESET_DROPDOWN_WIDTH = 160;         /* プリセットのドロップダウンの幅 / preset dropdown width */
    var PRESET_NAME_CHARS  = 20;             /* プリセット名の入力欄の文字数 / preset name field width */

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

    // アイコンのボタン（再利用パーツ） / Icon buttons (reusable)

    var ICON_BUTTON_SIZE = [24, 22]; /* 既定の大きさ / default size */
    var ICON_BUTTON_UI_DARK = isDarkUI();
    var ICON_BUTTON_COLOR     = ICON_BUTTON_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* 絵の色 / icon color */
    var ICON_BUTTON_DIM_COLOR = ICON_BUTTON_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の色 / color when disabled */

    /**
     * onDraw で絵を描くアイコンのボタンを追加する。押すと onClick を呼ぶ（無効の間は押せず、薄く描く）
     * @param {Group} parent - 追加先
     * @param {number[]} iconSize - [幅, 高さ]
     * @param {Function} drawIcon - 絵を描く関数 (iconGraphics, iconWidth, iconHeight, iconColor)
     * @returns {Group} アイコン
     */
    function addIconButton(parent, iconSize, drawIcon) {
        var iconButton = parent.add("group");
        iconButton.preferredSize = iconSize;
        iconButton.minimumSize = iconSize;
        iconButton.maximumSize = iconSize;
        iconButton.onDraw = function () {
            var iconColor = isIconButtonEnabledInTree(iconButton) ? ICON_BUTTON_COLOR : ICON_BUTTON_DIM_COLOR;
            drawIcon(iconButton.graphics, iconSize[0], iconSize[1], iconColor);
        };
        iconButton.addEventListener("mousedown", function () {
            if (!isIconButtonEnabledInTree(iconButton)) return;
            if (typeof iconButton.onClick === "function") iconButton.onClick();
        });
        return iconButton;
    }

    /**
     * アイコンのボタンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} iconButton - addIconButton() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setIconButtonEnabled(iconButton, isEnabled) {
        if (iconButton.enabled === isEnabled) return;
        iconButton.enabled = isEnabled;
        /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
        iconButton.hide();
        iconButton.show();
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isIconButtonEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 絵の座標の長方形の並びを、縦横比を保って中央に置いて塗る（ScriptUI は多角形を塗れないため、斜めも長方形の並びで描く）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} designSize - 絵の [幅, 高さ]
     * @param {number[][]} designRects - 長方形 [左, 上, 右, 下] の並び
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
        var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
        var originX = (iconWidth - designSize[0] * iconScale) / 2;
        var originY = (iconHeight - designSize[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < designRects.length; i++) {
            var designRect = designRects[i];
            if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
            iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 保存アイコン（トレイに下向きの矢印）を描く。1223×993 の絵。
     * 矢じりとトレイの V 字の切り欠きは細い長方形を並べ、トレイの四角い穴は塗らずに残す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSaveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designWidth = 1223;
        var trayHole = [78, 688, 230, 840]; /* トレイの四角い穴 [左, 上, 右, 下] / the tray's square hole */
        var sliceCount = 16;
        var designRects = [[535, 0, 688, 383]]; /* 矢印の軸 / arrow shaft */
        var k;

        /** 穴と重なる部分を除いて長方形を足す / add a rectangle minus the tray hole */
        function addTrayRect(left, top, right, bottom) {
            if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                designRects.push([left, top, right, bottom]);
                return;
            }
            designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
        }

        /* 下向きの矢じり（上の付け根から先端へ細っていく）/ the downward head */
        for (k = 0; k < sliceCount; k++) {
            var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
            var headTop = 383 + (703 - 383) * k / sliceCount;
            designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
        }
        /* トレイ：上辺に V 字の切り欠き（幅 458 から下へ細り、856 で閉じる）/ tray with a V notch along the top */
        for (k = 0; k < sliceCount; k++) {
            var notchTop = 612 + (856 - 612) * k / sliceCount;
            var notchBottom = notchTop + (856 - 612) / sliceCount;
            var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
            addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
            addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
        }
        addTrayRect(0, 856, designWidth, 993);
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [designWidth, 993], designRects, iconColor);
    }

    /**
     * 削除アイコン（ゴミ箱）を描く。756×825 の絵
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawTrashIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [756, 825], [
            [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],     /* 取っ手 / handle */
            [0, 137, 756, 207],                                               /* ふた / lid */
            [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825], /* 本体 / body */
            [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687] /* 縦の線 / ribs */
        ], iconColor);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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

    /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
    var STEPPER_POINTS_PER_UNIT = {
        "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
        "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
        "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
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
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
        var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
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
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
            if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = value * pointsPerUnit;
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldPointsPerUnit;
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
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
            title: { ja: "線と矢印を設定", en: "Set Stroke and Arrowheads" },
            presetSave: { ja: "プリセットを保存", en: "Save Preset" }
        },
        dropdown: {
            presetPlaceholder: { ja: "---", en: "---" }
        },
        panel: {
            stroke: { ja: "線", en: "Stroke" },
            arrowhead: { ja: "矢印", en: "Arrowheads" },
            dash: { ja: "破線", en: "Dashed Line" },
            calcMethod: { ja: "計算方法", en: "Calculation" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            presetName: { ja: "プリセット名", en: "Preset name" },
            strokeWidth: { ja: "線幅", en: "Weight" },
            strokeColor: { ja: "カラー", en: "Color" },
            strokeCap: { ja: "線端", en: "Cap" },
            cornerJoin: { ja: "角の形状", en: "Corner" },
            arrowScale: { ja: "倍率", en: "Scale" },
            segments: { ja: "分割数", en: "Segments" },
            gap: { ja: "間隔", en: "Gap" },
            dash: { ja: "線分", en: "Dash" }
        },
        unit: {
            point: { ja: "pt", en: "pt" },
            percent: { ja: "%", en: "%" }
        },
        radio: {
            gapToDash: { ja: "間隔 → 線分", en: "Gap → Dash" },
            dashToGap: { ja: "線分 → 間隔", en: "Dash → Gap" },
            noDash: { ja: "なし", en: "None" },
            dashed: { ja: "破線", en: "Dashed" },
            dotted: { ja: "ドット点線", en: "Dotted" },
            tipAtEnd: { ja: "パスの終点に配置", en: "At end of path" },
            tipBeyondEnd: { ja: "パスの終点から配置", en: "Beyond end of path" },
            keepDashLength: { ja: "線分の長さを保持", en: "Keep dash lengths" },
            adjustDashEnds: { ja: "両端を調整", en: "Adjust ends" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" },
            openStrokePanel: { ja: "「線」パネルを開く", en: "Open Stroke Panel" }
        },
        tooltip: {
            openStrokePanel: {
                ja: "ダイアログボックスを閉じて（設定は適用せずに）、Illustrator の［線］パネルを開きます。",
                en: "Closes the dialog without applying and opens Illustrator's Stroke panel."
            },
            preset: { ja: "保存した設定を読み込みます。", en: "Loads a saved set of settings." },
            presetSave: { ja: "今の設定に名前を付けて保存します。", en: "Saves the current settings under a name." },
            presetDelete: { ja: "選んでいるプリセットを削除します。", en: "Deletes the selected preset." },
            presetName: { ja: "保存する設定の名前です。同じ名前は上書きします。", en: "Name the settings are saved under. The same name is overwritten." },
            strokeWidth: { ja: "線の太さです。", en: "Weight of the stroke." },
            strokeWidthList: { ja: "よく使う線幅から選びます。", en: "Pick a common weight." },
            strokeColor: { ja: "クリックしてカラーピッカーで線の色を選びます。", en: "Click to pick the stroke color in the Color Picker." },
            roundCap: { ja: "option＋クリックで角の形状もラウンドにします。", en: "Option-click to set Round Join too." },
            favoriteArrow: {
                ja: "始点に付ける矢印です。倍率と先端位置はこの矢印に合わせた値に変わります。option＋クリックで始点と終点を入れ替え、⌘＋option＋クリックで［終点も同じ］を切り替えます。",
                en: "Arrowhead for the start of the path. The scale and tip alignment change to suit it. Option-click to swap the start and end; Cmd-Option-click to toggle Same at end."
            },
            otherArrow: {
                ja: "ほかの矢印をメニューから選びます。倍率は {scale}% に戻ります。",
                en: "Picks another arrowhead from the menu. The scale returns to {scale}%."
            },
            arrowScale: { ja: "矢印の大きさ（％）です。", en: "Size of the arrowhead, in percent." },
            sameEnd: {
                ja: "終点にも始点と同じ矢印を付けます。オフのときは終点の矢印を外します。",
                en: "Puts the same arrowhead on the end. When off, the end has no arrowhead."
            },
            swapEnds: {
                ja: "矢印を始点ではなく終点に付けます。",
                en: "Puts the arrowhead on the end of the path instead of the start."
            },
            noDash: { ja: "破線にしません（実線）。", en: "Leaves the line solid." },
            dashed: {
                ja: "破線にします。選ぶと、線幅から決めた分割数・間隔・線分が入ります。",
                en: "Makes a dashed line. Choosing it fills in Segments, Gap and Dash based on the stroke weight."
            },
            dotted: {
                ja: "点線にします。点の間隔は分割数から求め、線端は丸型に固定します。",
                en: "Makes a dotted line. The gap comes from Segments, and the cap is fixed to Round."
            },
            segments: {
                ja: "パスをいくつに分けるか。線分＋間隔の繰り返し回数になります。",
                en: "How many parts the path is divided into — the number of dash + gap cycles."
            },
            gap: { ja: "破線のすき間の長さです。", en: "Length of the empty space between dashes." },
            dash: { ja: "破線の線の長さです。", en: "Length of each dash." },
            gapToDash: { ja: "間隔を入力して線分の長さを求めます。", en: "Enter the gap; the dash length is calculated." },
            dashToGap: { ja: "線分の長さを入力して間隔を求めます。", en: "Enter the dash length; the gap is calculated." },
            adjustDashEnds: {
                ja: "オープンパスで、両端が線分で終わるように配分します（クローズパスでは使いません）。",
                en: "On an open path, distributes the dashes so both ends finish with a dash (unused for closed paths)."
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
        /* アクションに埋め込む Illustrator の表示名。矢印名はアイコンのツールチップとメニューの表示にも使う
           Illustrator labels embedded in the action; arrowhead names double as icon tooltips and menu labels */
        actionName: {
            noArrowhead: { ja: "[なし]", en: "[None]" },
            arrowheadPrefix: { ja: "矢印 ", en: "Arrow " },
            tipAtEnd: { ja: "パスの終点に配置", en: "Place Arrow Tip At End of Path" },
            tipBeyondEnd: { ja: "パスの終点から配置", en: "Extend Arrow Tip Beyond End of Path" },
            buttCap: { ja: "線端なし", en: "Butt Cap" },
            roundCap: { ja: "丸型線端", en: "Round Cap" },
            projectingCap: { ja: "突出線端", en: "Projecting Cap" },
            miterJoin: { ja: "マイター結合", en: "Miter Join" },
            roundJoin: { ja: "ラウンド結合", en: "Round Join" },
            bevelJoin: { ja: "ベベル結合", en: "Bevel Join" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select at least one object." },
            invalidWidth: { ja: "線幅には 0 以上の数値を入力してください。", en: "Enter a stroke weight of 0 or greater." },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." },
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the preset." }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete: { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        }
    };

    // =========================================
    // 設定の保存 / Settings store
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* プリセットは名前をキーにした集まり（既定値 {} で中身を問わず受け取る）。再起動しても残す
       Presets: a map keyed by name (the {} default accepts any content), kept across restarts */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "persistent", {
        /* 旧名 FavoriteArrow のときに保存したプリセットを読み継ぐ / carry over presets saved under the old name FavoriteArrow */
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/FavoriteArrowPresets.json");
        }
    });

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
     * 環境設定キーの単位を返す（Q/H の表示は使わないので、単位コード5は「Q」のまま）
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 一般の単位に合わせた線幅の初期値を返す（mm なら 0.25pt、px なら 1px）
     * @returns {number} 線幅の初期値（pt）
     */
    function getDefaultStrokeWidth() {
        var unitLabel = getUnitInfo("rulerType").label;
        return DEFAULT_STROKE_WIDTH_BY_UNIT.hasOwnProperty(unitLabel) ? DEFAULT_STROKE_WIDTH_BY_UNIT[unitLabel] : DEFAULT_STROKE_WIDTH;
    }

    // =========================================
    // 選択肢 / Options
    // =========================================

    /* アイコンの図形。size の枠の中に、長方形 rects [左, 上, 右, 下]、矢じり heads [付け根x, 先端x, 中心y, 半分の高さ]、
       折れ線 lines [[x, y], [x, y], …, 線幅]（省略可）で描く。縦横は同じ倍率で縮め、描く範囲の中央に置く
       Icon shapes: in a frame of the given size, rectangles [left, top, right, bottom], heads [base x, tip x, center y, half height]
       and optional polylines [[x, y], [x, y], ..., width]. Scaled uniformly and centered in the drawing area */
    var OPTION_ICON_SHAPES = {
        /* 両端を調整 / Adjust ends */
        keepLength: {
            size: [563, 450],
            rects: [
                [110, 53, 281, 110],                          /* 上の線分 / top dash */
                [395, 53, 509, 110], [452, 110, 509, 282],    /* 右上の L / top-right L */
                [53, 224, 110, 338], [53, 338, 167, 395],     /* 左下の L / bottom-left L */
                [281, 338, 452, 395]                          /* 下の線分 / bottom dash */
            ],
            heads: []
        },
        adjustEnds: {
            size: [563, 450],
            rects: [
                [53, 53, 225, 110], [53, 110, 111, 167],      /* 左上 / top-left */
                [338, 53, 510, 110], [452, 110, 510, 167],    /* 右上 / top-right */
                [53, 281, 111, 338], [53, 338, 225, 395],     /* 左下 / bottom-left */
                [452, 281, 510, 338], [338, 338, 510, 395]    /* 右下 / bottom-right */
            ],
            heads: []
        },
        /* 先端位置：上がパスの終点（T字）、下が矢印 / Tip alignment: the path end (T) above, the arrow below */
        beyondEnd: {
            size: [573, 420],
            rects: [
                [140, 114, 293, 139], [293, 87, 320, 166],     /* パスと終点 / path and its end */
                [155, 250, 335, 309]                          /* 矢印の軸（矢じりに食い込ませる）/ arrow shaft, overlapping the head */
            ],
            /* 細い長方形で描く矢じりは先端が細って短く見えるので、元の絵より少し右へ / sliced heads look short, so nudged right */
            heads: [[308, 451, 279, 77]]                      /* 終点から先へ出る矢じり / head beyond the end */
        },
        atEnd: {
            size: [573, 420],
            rects: [
                [172, 114, 440, 139], [440, 87, 467, 166],     /* パスと終点 / path and its end */
                [187, 250, 365, 309]                          /* 矢印の軸（矢じりに食い込ませる）/ arrow shaft, overlapping the head */
            ],
            heads: [[339, 482, 279, 77]]                      /* 先端が終点にそろう矢じり / head ending at the end */
        }
    };

    /* 線端・角の形状のアイコン（Illustrator の線パネルの絵に合わせる）。白い部分（パスと端点）は holes で地の色に抜く。
       丸い部分は discs [中心x, 中心y, 半径, 向き]（"left" は左半分の円、"topLeft" は左上の四分円）
       Cap and corner icons, after Illustrator's Stroke panel. The white path and end point are punched out with holes;
       round parts are discs [center x, center y, radius, side] ("left": left half, "topLeft": top-left quarter) */
    var STROKE_ICON_SHAPES = {
        butt: {
            size: [340, 360],
            rects: [[123, 20, 320, 340], [74, 119, 123, 241]],        /* 線と端点の枠 / stroke and end-point frame */
            heads: [],
            holes: [[99, 144, 172, 216], [172, 168, 320, 192]]        /* 端点とパス / end point and path */
        },
        round: {
            size: [340, 360],
            rects: [[185, 20, 320, 340]],
            heads: [],
            discs: [[185, 180, 160, "left"]],
            holes: [[124, 144, 197, 216], [197, 168, 320, 192]]
        },
        projecting: {
            size: [340, 360],
            rects: [[25, 20, 320, 340]],
            heads: [],
            holes: [[123, 144, 196, 216], [196, 168, 320, 192]]
        },
        miter: {
            size: [340, 334],
            rects: [[25, 20, 320, 314]],
            heads: [],
            holes: [[246, 241, 320, 314], [148, 191, 172, 314], [123, 118, 197, 191], [197, 142, 320, 167]]
        },
        roundJoin: {
            size: [340, 334],
            rects: [[159, 20, 320, 314], [25, 154, 159, 314]],
            heads: [],
            discs: [[159, 154, 134, "topLeft"]],
            holes: [[247, 241, 320, 314], [148, 191, 173, 314], [124, 118, 197, 191], [197, 142, 320, 167]]
        },
        bevel: {
            size: [340, 334],
            rects: [[148, 20, 320, 314], [25, 142, 50, 314], [50, 118, 74, 314], [74, 93, 99, 314], [99, 69, 123, 314], [123, 44, 148, 314]],
            heads: [],
            holes: [[246, 241, 320, 314], [148, 191, 172, 314], [123, 118, 196, 191], [196, 142, 320, 167]]
        }
    };

    /* よく使う矢印のアイコン（矢印の番号ごと。0 は［なし］で矢じりの無い線）。FAVORITE_ARROWS の番号にはすべて用意する
       Favorite arrowhead icons by number (0 is [None], a plain line). Every FAVORITE_ARROWS number needs one */
    var ARROW_ICON_SHAPES = {
        0: {
            size: [883, 383],
            rects: [[70, 179, 812, 205]],                     /* 矢じりの無い線 / plain line */
            heads: []
        },
        8: {
            size: [883, 383],
            rects: [[150, 154, 812, 230]],                    /* 軸（矢じりに食い込ませて継ぎ目を出さない）/ shaft, overlapping the head so no seam shows */
            heads: [[186, 70, 192, 116]]                      /* 塗りの矢じり / filled head */
        },
        11: {
            size: [883, 383],
            rects: [[100, 180, 812, 206]],                    /* 軸 / shaft */
            heads: [],
            /* 線の矢じり。折れ線1本で描いて先端を尖らせる / open head, one polyline so the tip comes to a point */
            lines: [[[191, 94], [73, 193], [191, 292], 24]]
        },
        13: {
            size: [883, 383],
            rects: [[73, 178, 812, 208]],                     /* 軸 / shaft */
            heads: [],
            /* 線の矢じり（折れ線）。端と先端は丸く（丸型線端・ラウンド結合の見た目）/ open head; round ends and tip */
            lines: [[[277, 66], [73, 193], [277, 320], 30]],
            discs: [[73, 193, 15, "full"], [277, 66, 15, "full"], [277, 320, 15, "full"], [812, 193, 15, "full"]]
        },
        21: {
            size: [883, 383],
            rects: [[96, 164, 783, 220]],                     /* 軸 / shaft */
            heads: [],
            discs: [[96, 192, 68, "full"], [783, 192, 28, "right"]] /* 端の丸と、反対側の丸い線端 / end dot and the round far end */
        },
        27: {
            size: [883, 383],
            rects: [[96, 178, 812, 204], [66, 109, 100, 274]], /* 軸と縦棒（縦棒は約2px。細すぎると1pxにつぶれる）/ shaft and bar, about 2 px so it does not collapse to 1 px */
            heads: []
        }
    };

    /* 両端を調整（オフ／オン）。アイコンで選ぶ / Adjust ends (off / on), picked with icons */
    var ADJUST_DASH_OPTIONS = [
        { key: "keepLength", label: "radio.keepDashLength" },
        { key: "adjustEnds", label: "radio.adjustDashEnds", tooltip: "tooltip.adjustDashEnds" }
    ];

    /* 先端位置・線端・角の形状の選択肢。value はアクションの enumerated 値
       線端・角の形状は表示名がそのまま説明になるので、helpTip にも actionName を使う
       Tip alignment, cap and join options; value is the action's enumerated value.
       The cap and join display names explain themselves, so they double as helpTips */
    var TIP_ALIGN_OPTIONS = [
        { key: "atEnd",     label: "radio.tipAtEnd",     actionName: "actionName.tipAtEnd",     value: 0 },
        { key: "beyondEnd", label: "radio.tipBeyondEnd", actionName: "actionName.tipBeyondEnd", value: 1 }
    ];
    var STROKE_CAP_OPTIONS = [
        { key: "butt",       label: "actionName.buttCap",       actionName: "actionName.buttCap",       value: 0, iconKey: "butt" },
        { key: "round",      label: "actionName.roundCap",      actionName: "actionName.roundCap",      value: 1, iconKey: "round", tooltip: "tooltip.roundCap" },
        { key: "projecting", label: "actionName.projectingCap", actionName: "actionName.projectingCap", value: 2, iconKey: "projecting" }
    ];
    var CORNER_JOIN_OPTIONS = [
        { key: "miter", label: "actionName.miterJoin", actionName: "actionName.miterJoin", value: 0, iconKey: "miter" },
        { key: "round", label: "actionName.roundJoin", actionName: "actionName.roundJoin", value: 1, iconKey: "roundJoin" },
        { key: "bevel", label: "actionName.bevelJoin", actionName: "actionName.bevelJoin", value: 2, iconKey: "bevel" }
    ];

    /**
     * 選択肢の一覧から key が一致するものを返す
     * @param {Object[]} optionDefinitions - 選択肢の一覧
     * @param {string} optionKey - 探す key
     * @returns {Object} 一致した選択肢。無ければ先頭
     */
    function findOptionByKey(optionDefinitions, optionKey) {
        for (var i = 0; i < optionDefinitions.length; i++) {
            if (optionDefinitions[i].key === optionKey) return optionDefinitions[i];
        }
        return optionDefinitions[0];
    }

    /**
     * よく使う矢印の名前の一覧を作る（FAVORITE_ARROWS の順）
     * @returns {string[]} 矢印名
     */
    function buildFavoriteArrowNames() {
        var favoriteArrowNames = [];
        for (var i = 0; i < FAVORITE_ARROWS.length; i++) {
            var arrowNumber = FAVORITE_ARROWS[i].number;
            favoriteArrowNames.push(arrowNumber === 0 ? getLabel("actionName.noArrowhead") : getLabel("actionName.arrowheadPrefix") + arrowNumber);
        }
        return favoriteArrowNames;
    }

    /**
     * 矢印の番号から FAVORITE_ARROWS 上の位置を返す
     * @param {number} arrowNumber - 矢印の番号（[なし] は 0）
     * @returns {number} FAVORITE_ARROWS の添字。よく使う矢印に無ければ -1
     */
    function findFavoriteArrowIndex(arrowNumber) {
        for (var i = 0; i < FAVORITE_ARROWS.length; i++) {
            if (FAVORITE_ARROWS[i].number === arrowNumber) return i;
        }
        return -1;
    }

    /**
     * 最初に選んでおく矢印の FAVORITE_ARROWS 上の位置を返す（DEFAULT_ARROW_NUMBER が無ければ 0）
     * @returns {number} FAVORITE_ARROWS の添字
     */
    function getDefaultFavoriteIndex() {
        return Math.max(0, findFavoriteArrowIndex(DEFAULT_ARROW_NUMBER));
    }

    /**
     * よく使う矢印に無い矢印名の一覧を作る（番号順）
     * @returns {string[]} ポップアップメニューに並べる矢印名
     */
    function buildOtherArrowNames() {
        var favoriteNumbers = {};
        for (var i = 0; i < FAVORITE_ARROWS.length; i++) {
            favoriteNumbers[FAVORITE_ARROWS[i].number] = true;
        }
        var otherArrowNames = [];
        for (var arrowNumber = 1; arrowNumber <= ARROW_COUNT; arrowNumber++) {
            if (!favoriteNumbers[arrowNumber]) otherArrowNames.push(getLabel("actionName.arrowheadPrefix") + arrowNumber);
        }
        return otherArrowNames;
    }

    // =========================================
    // 対象のパス / Target paths
    // =========================================

    /**
     * 線を持てるパスを、グループ・複合パスの中までたどって集める
     * @param {Array} pageItems - 選択やグループの中身
     * @returns {PathItem[]} 集めたパス
     */
    function collectStrokePaths(pageItems) {
        var strokePaths = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "PathItem") {
                strokePaths.push(pageItem);
            } else if (pageItem.typename === "CompoundPathItem") {
                strokePaths = strokePaths.concat(collectStrokePaths(pageItem.pathItems));
            } else if (pageItem.typename === "GroupItem") {
                strokePaths = strokePaths.concat(collectStrokePaths(pageItem.pageItems));
            }
        }
        return strokePaths;
    }

    /**
     * 破線の計算結果を表示するための、先頭のパスの長さと開閉を控える（undo で参照が切れるため値だけ持つ）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @returns {Object|null} { length, closed }。パスが無ければ null
     */
    function getFirstPathMetrics(strokePaths) {
        if (strokePaths.length === 0) return null;
        return { length: strokePaths[0].length, closed: strokePaths[0].closed };
    }

    /**
     * 選択しているパスのうち、線のある最初のパスの線幅を返す（線幅欄の初期値に使う）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @returns {number|null} 線幅（pt）。線のあるパスが無ければ null
     */
    function getSelectedStrokeWidth(strokePaths) {
        for (var i = 0; i < strokePaths.length; i++) {
            if (strokePaths[i].stroked) return strokePaths[i].strokeWidth;
        }
        return null;
    }

    // =========================================
    // 破線の計算 / Dash calculation
    // =========================================

    /**
     * 間隔から線分長と1周期（線分＋間隔）の長さを求める（DashGapCalculator から移植）
     * クローズパスと「両端を調整」OFFは 1周期＝全長÷分割数、
     * オープンパスで「両端を調整」ONは両端が線分で終わるように配分する。
     * @param {number} segments - 分割数
     * @param {number} gapPt - 間隔（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {Object} { dashPt:number, cyclePt:number }（計算できない場合は null）
     */
    function calcDashAndCyclePt(segments, gapPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(gapPt >= 0)) return null;

        if (isClosed || !adjustEnds) {
            var cyclePt = pathLen / segments;
            return { dashPt: cyclePt - gapPt, cyclePt: cyclePt };
        }

        /* 分割数＝線分の本数、間隔は（分割数−1）回 */
        if (segments === 1) return { dashPt: pathLen, cyclePt: pathLen + gapPt };

        var dashPt = (pathLen - gapPt * (segments - 1)) / segments;
        return { dashPt: dashPt, cyclePt: dashPt + gapPt };
    }

    /**
     * 線分長から間隔を逆算する（DashGapCalculator から移植）
     * @param {number} segments - 分割数
     * @param {number} dashPt - 線分長（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {number} 間隔（pt）。計算できない場合は null
     */
    function calcGapPtFromDashPt(segments, dashPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(dashPt >= 0)) return null;

        if (isClosed || !adjustEnds) return (pathLen / segments) - dashPt;

        if (segments === 1) return 0;
        return (pathLen - dashPt * segments) / (segments - 1);
    }

    /**
     * 分割数から 1 本のパスに付ける strokeDashes の値を求める（DashGapCalculator の「破線の計算」）。
     * 破線は計算方法に従い、ドット点線は点（長さ0）を保って間隔を求める
     * @param {Object} dashCalc - { style:"dashed"/"dotted", mode:"gapToDash"/"dashToGap", segments, gapPt, dashPt, adjustEnds }
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @returns {number[]|null} [線分, 間隔]（pt）。計算できない場合は null
     */
    function calcDashArray(dashCalc, pathLen, isClosed) {
        var segments = dashCalc.segments;
        var adjustEnds = dashCalc.adjustEnds;
        if (!(segments >= 1)) return null;

        if (dashCalc.style === "dotted") {
            /* 両端に点を置くなら最低2つ / Dots at both ends need at least two */
            var dotCount = (!isClosed && adjustEnds) ? Math.max(2, segments) : segments;
            var solvedGapPt = calcGapPtFromDashPt(dotCount, 0, pathLen, isClosed, adjustEnds);
            if (solvedGapPt == null || solvedGapPt <= 0) return null;
            /* 両端を調整では最後の点がちょうど終点に来る。長さ0の点は終点をわずかでも越えると描かれないので、
               間隔を 0.001pt 単位で切り捨てて終点の手前に収める
               With adjusted ends the last dot lands exactly on the end; a zero-length dot past the end by any
               rounding error is not drawn, so truncate the gap to 0.001 pt to keep it inside */
            if (!isClosed && adjustEnds) solvedGapPt = Math.floor(solvedGapPt * DOT_GAP_PRECISION) / DOT_GAP_PRECISION;
            return [0, solvedGapPt];
        }

        /* 線分→間隔：オープンパスで両端を調整、分割数1ならパス全長を線分にする / Dash to gap */
        if (dashCalc.mode === "dashToGap") {
            var dashPt = (!isClosed && adjustEnds && segments === 1) ? pathLen : dashCalc.dashPt;
            if (!(dashPt >= 0)) return null;
            var gapFromDash = calcGapPtFromDashPt(segments, dashPt, pathLen, isClosed, adjustEnds);
            if (gapFromDash == null || gapFromDash < 0) return null;
            return [dashPt, gapFromDash];
        }

        /* 間隔→線分 / Gap to dash */
        if (!(dashCalc.gapPt >= 0)) return null;
        var dashCycle = calcDashAndCyclePt(segments, dashCalc.gapPt, pathLen, isClosed, adjustEnds);
        if (!dashCycle || dashCycle.dashPt <= 0) return null;
        return [dashCycle.dashPt, dashCalc.gapPt];
    }

    /**
     * 線幅比のパターンに近い分割数を求める（オンにしたときの初期値）
     * @param {string} dashStyle - "dashed" / "dotted"
     * @param {number} strokeWidth - 線幅（pt）
     * @param {Object} pathMetrics - getFirstPathMetrics() の戻り値
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {number} 分割数（1以上の整数）
     */
    function estimateSegments(dashStyle, strokeWidth, pathMetrics, adjustEnds) {
        var dashPattern = DASH_PATTERNS[dashStyle];
        var cyclePt = (dashPattern.dash + dashPattern.gap) * strokeWidth;
        if (!(cyclePt > 0) || !(pathMetrics.length > 0)) return 1;
        /* 両端を調整：線分 n 本と間隔 n−1 個 / Adjusted ends: n dashes and n - 1 gaps */
        var gapPt = dashPattern.gap * strokeWidth;
        var estimate = (!pathMetrics.closed && adjustEnds) ? (pathMetrics.length + gapPt) / cyclePt : pathMetrics.length / cyclePt;
        return Math.max(1, Math.round(estimate));
    }

    /**
     * 破線を DOM で設定する。アクションの後に呼ぶ。破線が［なし］なら破線を外して実線にする
     * （アクションだけでは元の破線が残ることがあるため、DOM で明示的に外す）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @param {Object|null} dashCalc - readDashCalc() の戻り値。null なら実線にする
     * @returns {void}
     */
    function applyDashStyle(strokePaths, dashCalc) {
        var i;
        if (!dashCalc) {
            for (i = 0; i < strokePaths.length; i++) {
                strokePaths[i].strokeDashes = [];
                strokePaths[i].strokeDashOffset = 0;
            }
            return;
        }
        for (i = 0; i < strokePaths.length; i++) {
            var dashArray = calcDashArray(dashCalc, strokePaths[i].length, strokePaths[i].closed);
            if (!dashArray) continue;
            strokePaths[i].strokeDashOffset = 0;
            strokePaths[i].strokeDashes = dashArray;
        }
    }

    /**
     * 線の設定をまとめて適用する（アクション → 破線の順。アクションが破線を解除するため）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @param {Object} strokeSettings - readDialogSettings() の戻り値
     * @returns {void}
     */
    function applyStrokeSettings(strokePaths, strokeSettings) {
        /* 元の色はアクションの前に控える（線の無いパスはアクションで線が付き、色が変わるため）/ note the colors before the action adds strokes */
        var originalColors = collectOriginalColors(strokePaths);
        playStrokeAction(strokeSettings);
        applyDashStyle(strokePaths, strokeSettings.dashCalc);
        applyStrokeColor(strokePaths, strokeSettings.strokeColor, originalColors);
    }

    /**
     * パスごとの元の色を控える（線があれば線の色、線が無ければ塗りの色）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @returns {Array} 色の並び（色が無ければ null）
     */
    function collectOriginalColors(strokePaths) {
        var originalColors = [];
        for (var i = 0; i < strokePaths.length; i++) originalColors.push(getPathColor(strokePaths[i]));
        return originalColors;
    }

    /**
     * パスの色を返す（線があれば線の色、線が無ければ塗りの色）
     * @param {PathItem} strokePath - パス
     * @returns {Color|null} 色。線も塗りも無ければ null
     */
    function getPathColor(strokePath) {
        if (strokePath.stroked) return strokePath.strokeColor;
        if (strokePath.filled) return strokePath.fillColor;
        return null;
    }

    /**
     * 線の色を DOM で設定する。カラーを変えていればその色を、変えていなければパスごとの元の色を線の色にする
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @param {Color|null} strokeColor - カラーで選んだ色。変えていなければ null
     * @param {Array} originalColors - collectOriginalColors() の戻り値
     * @returns {void}
     */
    function applyStrokeColor(strokePaths, strokeColor, originalColors) {
        for (var i = 0; i < strokePaths.length; i++) {
            var pathColor = strokeColor || originalColors[i];
            if (pathColor) strokePaths[i].strokeColor = pathColor;
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューを制御する。線幅は DOM で即時プレビュー（元の値を控えて書き戻すので undo に頼らない）、
     * 矢印は DOM に無いのでアクション実行＋app.undo() でプレビューする
     * @returns {Object} previewStrokeWidth / previewSettings / clearAction / reset を持つオブジェクト
     */
    function createPreviewController() {
        var strokePaths = [];
        var originalStates = [];
        var isActionPreviewApplied = false;

        /* 現在の線設定を控える / Capture the current stroke state */
        function captureStrokeState() {
            strokePaths = collectStrokePaths(app.activeDocument.selection);
            originalStates = [];
            for (var i = 0; i < strokePaths.length; i++) {
                var strokePath = strokePaths[i];
                originalStates.push({
                    stroked: strokePath.stroked,
                    strokeWidth: strokePath.strokeWidth,
                    strokeColor: strokePath.strokeColor,
                    strokeDashes: strokePath.strokeDashes,
                    strokeDashOffset: strokePath.strokeDashOffset
                });
            }
        }

        /* 控えた線設定を書き戻す / Restore the captured stroke state */
        function restoreStrokeState() {
            for (var i = 0; i < strokePaths.length; i++) {
                var originalState = originalStates[i];
                strokePaths[i].stroked = originalState.stroked;
                if (originalState.stroked) {
                    strokePaths[i].strokeWidth = originalState.strokeWidth;
                    strokePaths[i].strokeColor = originalState.strokeColor;
                }
                strokePaths[i].strokeDashes = originalState.strokeDashes;
                strokePaths[i].strokeDashOffset = originalState.strokeDashOffset;
            }
        }

        /* アクションによるプレビューを undo で取り消す / Undo the action based preview */
        function clearActionPreview() {
            if (!isActionPreviewApplied) return;
            /* 取り消す履歴が無いと例外になる / app.undo() throws when there is nothing to undo */
            try { app.undo(); } catch (e) {}
            isActionPreviewApplied = false;
            /* undo でオブジェクト参照が無効になるため取り直す / References die on undo, re-capture */
            captureStrokeState();
        }

        captureStrokeState();

        return {
            /* 線幅のみ DOM で即時プレビュー / Preview stroke width only, via the DOM */
            previewStrokeWidth: function (strokeWidth) {
                clearActionPreview();
                if (isNaN(strokeWidth) || strokeWidth < 0) return;
                for (var i = 0; i < strokePaths.length; i++) {
                    strokePaths[i].stroked = true;
                    strokePaths[i].strokeWidth = strokeWidth;
                }
                app.redraw();
            },

            /* 矢印を含む全設定をアクションで 1 回だけプレビュー / Preview all settings via the action */
            previewSettings: function (strokeSettings) {
                clearActionPreview();
                restoreStrokeState();
                applyStrokeSettings(strokePaths, strokeSettings);
                isActionPreviewApplied = true;
                app.redraw();
            },

            /* アクションによるプレビューだけを取り消す / Clear only the action based preview */
            clearAction: function () {
                if (!isActionPreviewApplied) return;
                clearActionPreview();
                app.redraw();
            },

            /* すべてのプレビューを取り消して元の状態に戻す / Revert everything */
            reset: function () {
                clearActionPreview();
                restoreStrokeState();
                app.redraw();
            }
        };
    }

    // =========================================
    // ダイアログの部品 / Dialog parts
    // =========================================

    /**
     * 線の色の色見本を追加する（onDraw で塗る）。クリックで Illustrator 標準のカラーピッカーを開く。
     * 色は .strokeColor に持ち、カラーピッカーで変えたら .isChanged を true にして .onColorChange を呼ぶ
     * @param {Group} parent - 追加先
     * @returns {Group} 色見本
     */
    function addColorSwatch(parent) {
        var colorSwatch = parent.add("group");
        colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        colorSwatch.minimumSize = COLOR_SWATCH_SIZE;
        colorSwatch.maximumSize = COLOR_SWATCH_SIZE;
        colorSwatch.helpTip = getLabel("tooltip.strokeColor");
        colorSwatch.strokeColor = createBlackColor();
        colorSwatch.isChanged = false;
        colorSwatch.onDraw = function () { drawColorSwatch(colorSwatch); };
        colorSwatch.addEventListener("mousedown", function () {
            /* キャンセルでは渡した色がそのまま返るので、中身を比べて変化を見る / Cancel returns the passed color, so compare */
            var pickedColor = app.showColorPicker(colorSwatch.strokeColor);
            if (!pickedColor || getColorKey(pickedColor) === getColorKey(colorSwatch.strokeColor)) return;
            setSwatchColor(colorSwatch, pickedColor);
            colorSwatch.isChanged = true;
            if (typeof colorSwatch.onColorChange === "function") colorSwatch.onColorChange();
        });
        return colorSwatch;
    }

    /**
     * 色見本の色を変えて描き直す
     * @param {Group} colorSwatch - addColorSwatch() で作った色見本
     * @param {Color} swatchColor - 色
     * @returns {void}
     */
    function setSwatchColor(colorSwatch, swatchColor) {
        colorSwatch.strokeColor = swatchColor;
        redrawStepperGroup(colorSwatch);
    }

    /**
     * 色見本を塗る（色が無い・表せないときは白地に赤の斜線）
     * @param {Group} colorSwatch - addColorSwatch() で作った色見本
     * @returns {void}
     */
    function drawColorSwatch(colorSwatch) {
        var swatchGraphics = colorSwatch.graphics;
        var swatchWidth = COLOR_SWATCH_SIZE[0];
        var swatchHeight = COLOR_SWATCH_SIZE[1];
        var displayRgb = getDisplayRgb(colorSwatch.strokeColor);
        swatchGraphics.newPath();
        swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
        swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, displayRgb ? displayRgb.concat([1]) : [1, 1, 1, 1]));
        if (!displayRgb) {
            swatchGraphics.newPath();
            swatchGraphics.moveTo(1, swatchHeight - 1);
            swatchGraphics.lineTo(swatchWidth - 1, 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.9, 0.1, 0.1, 1], 1.5));
        }
        swatchGraphics.newPath();
        swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
        swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
    }

    /**
     * 色見本に塗る RGB（0〜1）を求める。CMYK・グレーは簡易換算、スポットは元の色で、それ以外は null
     * @param {Color} sourceColor - 色
     * @returns {number[]|null} [r, g, b]。色が無い・グラデーション・パターンなら null
     */
    function getDisplayRgb(sourceColor) {
        if (!sourceColor) return null;
        switch (sourceColor.typename) {
            case "RGBColor":
                return [sourceColor.red / 255, sourceColor.green / 255, sourceColor.blue / 255];
            case "CMYKColor":
                var blackRatio = sourceColor.black / 100;
                return [(1 - sourceColor.cyan / 100) * (1 - blackRatio), (1 - sourceColor.magenta / 100) * (1 - blackRatio), (1 - sourceColor.yellow / 100) * (1 - blackRatio)];
            case "GrayColor":
                var grayLevel = 1 - sourceColor.gray / 100;
                return [grayLevel, grayLevel, grayLevel];
            case "SpotColor":
                return getDisplayRgb(sourceColor.spot.color);
            default:
                return null;
        }
    }

    /**
     * 色を比べるための文字列にする（カラーピッカーのキャンセルを見分けるため）
     * @param {Color} sourceColor - 色
     * @returns {string} 型と値をつないだ文字列
     */
    function getColorKey(sourceColor) {
        switch (sourceColor.typename) {
            case "RGBColor": return ["RGB", sourceColor.red, sourceColor.green, sourceColor.blue].join(",");
            case "CMYKColor": return ["CMYK", sourceColor.cyan, sourceColor.magenta, sourceColor.yellow, sourceColor.black].join(",");
            case "GrayColor": return ["Gray", sourceColor.gray].join(",");
            case "SpotColor": return ["Spot", sourceColor.spot.name, sourceColor.tint].join(",");
            default: return sourceColor.typename;
        }
    }

    /**
     * 黒（グレー100%）を作る。ドキュメントのカラーモードを問わず黒になる
     * @returns {GrayColor} 黒
     */
    function createBlackColor() {
        var blackColor = new GrayColor();
        blackColor.gray = 100;
        return blackColor;
    }

    /**
     * 選択している最初のパスの色を返す（色見本の初期値に使う。線が無ければ塗りの色）
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @returns {Color|null} 色。色のあるパスが無ければ null
     */
    function getSelectedStrokeColor(strokePaths) {
        for (var i = 0; i < strokePaths.length; i++) {
            var pathColor = getPathColor(strokePaths[i]);
            if (pathColor) return pathColor;
        }
        return null;
    }

    /**
     * 線幅欄のすぐ右に、よく使う線幅を選ぶ▼ボタン（onDraw で描く）を足す。
     * ScriptUI にはコンボボックスが無く、ドロップダウンリストを開く処理も呼べないので、
     * 押したら欄の下に小さなリストのダイアログボックスを出す
     * @param {EditText} strokeWidthInput - addNumberField() で作った線幅欄
     * @returns {Group} ▼ボタン（.onPick(線幅) を呼び出し側で付ける）
     */
    function addStrokeWidthPopupButton(strokeWidthInput) {
        /* 入力欄と隙間0で突き合わせる（∧∨と入力欄が入っているグループに足す）/ butt it against the field */
        var popupButton = strokeWidthInput.parent.add("group");
        var buttonSize = [WIDTH_POPUP_BUTTON_WIDTH, strokeWidthInput.preferredSize.height > 0 ? strokeWidthInput.preferredSize.height : WIDTH_POPUP_BUTTON_HEIGHT];
        popupButton.preferredSize = buttonSize;
        popupButton.minimumSize = buttonSize;
        popupButton.maximumSize = buttonSize;
        popupButton.helpTip = getLabel("tooltip.strokeWidthList");
        popupButton.onDraw = function () {
            var buttonGraphics = popupButton.graphics;
            var buttonWidth = popupButton.size ? popupButton.size[0] : buttonSize[0];
            var buttonHeight = popupButton.size ? popupButton.size[1] : buttonSize[1];
            var isDark = isDarkUI();
            var frameColor = isDark ? [1, 1, 1, 0.25] : [0, 0, 0, 0.25];
            var chevronColor = isDark ? [1, 1, 1, 0.85] : [0, 0, 0, 0.65];
            /* 枠（左も描いて入力欄との境目にする）/ full frame; the left edge marks the border with the field */
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonWidth - 1, buttonHeight - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, frameColor, 1));
            /* 下向きの山形 / downward chevron */
            var centerX = buttonWidth / 2;
            var centerY = buttonHeight / 2;
            buttonGraphics.newPath();
            buttonGraphics.moveTo(centerX - 4, centerY - 2);
            buttonGraphics.lineTo(centerX, centerY + 2);
            buttonGraphics.lineTo(centerX + 4, centerY - 2);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, chevronColor, 1.5));
        };
        popupButton.addEventListener("mousedown", function () {
            var pickedIndex = showValuePopup(strokeWidthInput, STROKE_WIDTH_PRESETS, getLabel("unit.point"), parseFloat(strokeWidthInput.text));
            if (pickedIndex >= 0 && typeof popupButton.onPick === "function") popupButton.onPick(STROKE_WIDTH_PRESETS[pickedIndex]);
        });
        return popupButton;
    }

    /**
     * コントロールの左上の画面上の位置を求める（親をたどって bounds を足し、ウィンドウの中身の位置を足す）
     * @param {Object} control - コントロール
     * @returns {number[]} [x, y]
     */
    function getScreenPosition(control) {
        var screenX = 0;
        var screenY = 0;
        for (var node = control; node && node.parent; node = node.parent) {
            screenX += node.bounds[0];
            screenY += node.bounds[1];
        }
        screenX += control.window.bounds[0];
        screenY += control.window.bounds[1];
        return [screenX, screenY];
    }

    /**
     * 値のリストを、基準のコントロールの下に枠なしの小さなダイアログボックスで出し、選んだ位置を返す
     * （クリックで決定、Esc や外側をクリックで閉じる）
     * @param {Object} anchorControl - この下に出す
     * @param {number[]} values - 並べる値
     * @param {string} unitLabel - 値の後ろに付ける単位
     * @param {number} currentValue - 今の値（同じ値を選んだ状態で開く）
     * @returns {number} 選んだ値の位置。選ばずに閉じたら -1
     */
    function showValuePopup(anchorControl, values, unitLabel, currentValue) {
        var popupDialog = new Window("dialog", undefined, undefined, { borderless: true });
        popupDialog.margins = 2;
        var valueItems = [];
        for (var i = 0; i < values.length; i++) valueItems.push(values[i] + " " + unitLabel);
        var valueList = popupDialog.add("listbox", undefined, valueItems);
        valueList.preferredSize = [Math.max(anchorControl.size ? anchorControl.size[0] : 0, WIDTH_POPUP_LIST_SIZE[0]), WIDTH_POPUP_LIST_SIZE[1]];
        for (var j = 0; j < values.length; j++) {
            if (Math.abs(values[j] - currentValue) < 0.001) valueList.selection = j;
        }
        var pickedIndex = -1;
        valueList.onChange = function () {
            if (!valueList.selection) return;
            pickedIndex = valueList.selection.index;
            popupDialog.close(1);
        };
        /* 外側をクリックしてフォーカスが外れたら閉じる / close when focus moves away */
        popupDialog.onDeactivate = function () { popupDialog.close(2); };
        var anchorPosition = getScreenPosition(anchorControl);
        popupDialog.location = [anchorPosition[0], anchorPosition[1] + (anchorControl.size ? anchorControl.size[1] : 0) + 2];
        popupDialog.show();
        return pickedIndex;
    }

    /* 項目名の幅をそろえる（_templates/AlignLabelWidths.jsx から）/ align label widths (from _templates/AlignLabelWidths.jsx) */
    /**
     * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
     * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
     * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
     * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
     * @returns {number} 揃えた後の共通ラベル幅（px）。
     */
    function alignLabelWidths(labelControls, justify) {
        if (!labelControls || !labelControls.length) return 0;
        justify = justify || 'right';

        /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
            var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }

        /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = justify;
        }

        return maxLabelWidth;
    }

    /**
     * 上に余白を取った入れ子のパネルを追加する（余白はパネルを包むグループの margins で取る）
     * @param {Panel} parent - 追加先のパネル
     * @param {string} titleKey - パネル名の LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addSubPanel(parent, titleKey) {
        var subPanelGroup = parent.add("group");
        subPanelGroup.orientation = "column";
        subPanelGroup.alignChildren = ["fill", "top"];
        subPanelGroup.alignment = ["fill", "top"];
        subPanelGroup.margins = [0, SUB_PANEL_TOP_MARGIN, 0, 0];
        var subPanel = subPanelGroup.add("panel", undefined, getLabel(titleKey));
        setupPanel(subPanel, 6);
        return subPanel;
    }

    /**
     * 横幅いっぱいの区切り線（高さ1の panel）を追加する。区切り線は setupPanel を通さない
     * @param {Panel|Group} parent - 追加先
     * @returns {Panel} 区切り線
     */
    function addSeparator(parent) {
        var separatorLine = parent.add("panel");
        separatorLine.alignment = ["fill", "top"];
        separatorLine.minimumSize.height = 1;
        separatorLine.maximumSize.height = 1;
        return separatorLine;
    }

    /**
     * 右揃えの項目名を持つ行を追加する。項目名の幅は文字なりのまま（alignLabelWidths() でそろえる）
     * @param {Panel|Group} parent - 追加先
     * @param {string} labelKey - 項目名の LABELS パス
     * @returns {Group} 追加した行（項目名は .rowLabel）
     */
    function addLabeledRow(parent, labelKey) {
        var labeledRow = parent.add("group");
        setupRow(labeledRow, "left", LABEL_GAP);
        var rowLabel = labeledRow.add("statictext", undefined, labelText(labelKey));
        rowLabel.justify = "right";
        labeledRow.rowLabel = rowLabel;
        return labeledRow;
    }

    /**
     * 幅いっぱいに広げないラジオボタンを追加する
     * @param {Panel|Group} parent - 追加先
     * @param {string} radioText - 表示名
     * @param {string} tooltipText - helpTip
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addOptionRadio(parent, radioText, tooltipText) {
        var optionRadio = parent.add("radiobutton", undefined, radioText);
        optionRadio.helpTip = tooltipText;
        optionRadio.alignment = "left";
        return optionRadio;
    }

    /**
     * ラジオボタンの代わりに、onDraw で描いたアイコンを横に並べる。
     * 各アイコンはラジオボタンと同じく value / optionDefinition / onClick を持つ（getCheckedOption・checkOptionByKey でそのまま扱える）
     * @param {Panel|Group} parent - 追加先
     * @param {Object[]} optionDefinitions - 選択肢（key / label / tooltip（省略可）。図形は iconKey があれば STROKE_ICON_SHAPES、無ければ OPTION_ICON_SHAPES[key]）
     * @param {string} selectedKey - 初期選択の key
     * @param {number[]} iconSize - アイコン1つの [幅, 高さ]
     * @param {number} [contentInset] - 枠と図形の間（px、省略時は 0）
     * @returns {Group[]} 追加したアイコン
     */
    function addOptionIcons(parent, optionDefinitions, selectedKey, iconSize, contentInset) {
        /* 隙間0で突き合わせ、2つ目以降は左の枠を描かずに1本の境界線を共有する
           Butted together; later icons skip their left edge so neighbors share one border */
        var iconRow = parent.add("group");
        setupRow(iconRow, "left", 0);
        var optionIcons = [];
        for (var i = 0; i < optionDefinitions.length; i++) {
            optionIcons.push(addOptionIcon(iconRow, optionDefinitions[i], optionIcons, iconSize));
            optionIcons[i].value = (optionDefinitions[i].key === selectedKey);
            optionIcons[i].hasLeftEdge = (i === 0);
            optionIcons[i].contentInset = contentInset || 0;
        }
        return optionIcons;
    }

    /**
     * アイコンの組の有効／無効を切り替えて描き直す（自作描画は自動でディムにならない）
     * @param {Group[]} optionIcons - addOptionIcons() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setOptionIconsEnabled(optionIcons, isEnabled) {
        for (var i = 0; i < optionIcons.length; i++) {
            optionIcons[i].enabled = isEnabled;
            redrawStepperGroup(optionIcons[i]);
        }
    }

    /**
     * 選択肢のアイコンを1つ作る。押すと同じ組のほかのアイコンを外し、onClick を呼ぶ
     * @param {Group} iconRow - 追加先の行
     * @param {Object} optionDefinition - 選択肢
     * @param {Group[]} optionIcons - 同じ組のアイコン（排他にする）
     * @param {number[]} iconSize - [幅, 高さ]
     * @returns {Group} アイコン
     */
    function addOptionIcon(iconRow, optionDefinition, optionIcons, iconSize) {
        var optionIcon = iconRow.add("group");
        optionIcon.preferredSize = iconSize;
        optionIcon.minimumSize = iconSize;
        optionIcon.maximumSize = iconSize;
        optionIcon.iconSize = iconSize;
        optionIcon.helpTip = getLabel(optionDefinition.label) + (optionDefinition.tooltip ? "\n" + getLabel(optionDefinition.tooltip) : "");
        optionIcon.optionDefinition = optionDefinition;
        optionIcon.value = false;
        optionIcon.onDraw = function () { drawOptionIcon(optionIcon); };
        optionIcon.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(optionIcon)) return;
            for (var i = 0; i < optionIcons.length; i++) {
                optionIcons[i].value = (optionIcons[i] === optionIcon);
                redrawStepperGroup(optionIcons[i]);
            }
            if (typeof optionIcon.onClick === "function") optionIcon.onClick();
        });
        return optionIcon;
    }

    /**
     * 選択肢のアイコンを描く（選択中は灰色の地、そうでなければ白の地。ダークUIでは明暗を反転）
     * @param {Group} optionIcon - addOptionIcon() で作ったアイコン
     * @returns {void}
     */
    function drawOptionIcon(optionIcon) {
        var iconGraphics = optionIcon.graphics;
        var isDark = isDarkUI();
        var isDimmed = !isStepperEnabledInTree(optionIcon);
        var inkLevel = isDark ? 0.65 : 0.45; /* 図形と枠は真っ黒・真っ白より抑える / keep the ink softer than pure black or white */
        var groundLevel = optionIcon.value ? (isDark ? 0.45 : 0.7) : (isDark ? 0.2 : 1);
        var inkColor = [inkLevel, inkLevel, inkLevel, isDimmed ? 0.4 : 1];
        var groundColor = [groundLevel, groundLevel, groundLevel, isDimmed ? 0.4 : 1];
        var iconWidth = optionIcon.iconSize[0];
        var iconHeight = optionIcon.iconSize[1];

        /* 地と枠。左隣と接するアイコンは左の枠を描かない（隣の右の枠と共有）
           Ground and frame; an icon next to another skips its left edge (shared with the neighbor's right edge) */
        var groundLeft = optionIcon.hasLeftEdge ? 0 : -0.5;
        iconGraphics.newPath();
        iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
        iconGraphics.newPath();
        if (optionIcon.hasLeftEdge) {
            iconGraphics.moveTo(0.5, iconHeight - 0.5);
            iconGraphics.lineTo(0.5, 0.5);
        } else {
            iconGraphics.moveTo(groundLeft, 0.5);
        }
        iconGraphics.lineTo(iconWidth - 0.5, 0.5);
        iconGraphics.lineTo(iconWidth - 0.5, iconHeight - 0.5);
        iconGraphics.lineTo(optionIcon.hasLeftEdge ? 0.5 : groundLeft, iconHeight - 0.5);
        iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, 1));

        /* 図形。枠の内側に contentInset の余白を取って縮める / shapes, shrunk inside the inset */
        var optionDefinition = optionIcon.optionDefinition;
        var iconShape = optionIcon.iconShape || (optionDefinition.iconKey ? STROKE_ICON_SHAPES[optionDefinition.iconKey] : OPTION_ICON_SHAPES[optionDefinition.key]);
        if (!iconShape) return;
        /* 両端に付けるときは左右反転した図形を重ねる（よく使う矢印のアイコン）/ overlay a mirrored copy for both ends */
        if (optionIcon.showBothEnds) iconShape = buildBothEndsShape(iconShape);
        var inset = optionIcon.contentInset || 0;
        drawIconShapes(iconGraphics, iconShape, [inset, inset, iconWidth - inset * 2, iconHeight - inset * 2], inkColor, optionIcon.shapeScale, groundColor);
    }

    /**
     * 図形に、左右反転した図形を足す（片側の矢印を両端の矢印にする）
     * @param {Object} iconShape - { size, rects, heads, lines（省略可） }
     * @returns {Object} 両端の図形
     */
    function buildBothEndsShape(iconShape) {
        var shapeWidth = iconShape.size[0];
        var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [], discs: [] };
        var mirroredSides = { left: "right", right: "left", full: "full" };
        var i, j;
        for (i = 0; i < iconShape.rects.length; i++) {
            var shapeRect = iconShape.rects[i];
            bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
        }
        for (i = 0; i < iconShape.heads.length; i++) {
            var headShape = iconShape.heads[i];
            bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
        }
        var shapeDiscs = iconShape.discs || [];
        for (i = 0; i < shapeDiscs.length; i++) {
            var discShape = shapeDiscs[i];
            bothEndsShape.discs.push(discShape, [shapeWidth - discShape[0], discShape[1], discShape[2], mirroredSides[discShape[3]] || discShape[3]]);
        }
        var shapeLines = iconShape.lines || [];
        for (i = 0; i < shapeLines.length; i++) {
            var shapeLine = shapeLines[i];
            var mirroredLine = [];
            for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
            mirroredLine.push(shapeLine[shapeLine.length - 1]);
            bothEndsShape.lines.push(shapeLine, mirroredLine);
        }
        return bothEndsShape;
    }

    /**
     * アイコンの図形（長方形・矢じり・線）を、指定の範囲に合わせて拡大縮小して描く
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {Object} iconShape - { size, rects, heads, discs / holes / lines（省略可） }
     * @param {number[]} drawArea - 描く範囲 [左, 上, 幅, 高さ]
     * @param {number[]} inkColor - [r, g, b, a]
     * @param {number} [fixedScale] - 絵の倍率を固定するとき（省略時は範囲に収まる倍率）
     * @param {number[]} [groundColor] - holes を抜く地の色 [r, g, b, a]（holes があるときに使う）
     * @returns {void}
     */
    function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
        /* 縦横比を保って縮め、中央に置く（縦横別の倍率だと矢じりの角度が変わる）
           Uniform scale, centered; separate x / y scales would change the head angles */
        var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
        var scaleX = iconScale;
        var scaleY = iconScale;
        var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
        var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < iconShape.rects.length; i++) {
            var shapeRect = iconShape.rects[i];
            var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
            var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
            var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
            var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
            iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
        }
        for (var j = 0; j < iconShape.heads.length; j++) {
            var headShape = iconShape.heads[j];
            addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
        }
        var shapeDiscs = iconShape.discs || [];
        for (var d = 0; d < shapeDiscs.length; d++) {
            var discShape = shapeDiscs[d];
            addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));

        /* 白い部分（パスと端点）は地の色で抜く / punch out the white parts with the ground color */
        var shapeHoles = iconShape.holes || [];
        if (shapeHoles.length > 0 && groundColor) {
            iconGraphics.newPath();
            for (var h = 0; h < shapeHoles.length; h++) {
                var holeRect = shapeHoles[h];
                var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                var holeTop = Math.round(originY + holeRect[1] * scaleY);
                iconGraphics.rectPath(holeLeft, holeTop,
                    Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
            }
            iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
        }

        var shapeLines = iconShape.lines || [];
        for (var k = 0; k < shapeLines.length; k++) {
            var shapeLine = shapeLines[k];
            var lineWidth = shapeLine[shapeLine.length - 1];
            /* 折れ線は1本ずつ newPath から描く（続けると前の線に積み重なる）/ one path per polyline */
            iconGraphics.newPath();
            for (var m = 0; m < shapeLine.length - 1; m++) {
                var pointX = originX + shapeLine[m][0] * scaleX;
                var pointY = originY + shapeLine[m][1] * scaleY;
                if (m === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
        }
    }

    /**
     * よく使う矢印のアイコンを1つ追加する（ラジオボタンの代わり。選択中は灰色の地）。
     * ラジオボタンと同じく value / onClick を持ち、矢印名は .arrowName に持つ。
     * value を変えたあとの描き直しは redrawFavoriteArrowIcons() で行う
     * @param {Group} parent - 追加先
     * @param {Object} favoriteArrow - FAVORITE_ARROWS の要素
     * @param {string} arrowName - 矢印名
     * @returns {Group} アイコン
     */
    function addFavoriteArrowIcon(parent, favoriteArrow, arrowName) {
        var arrowIcon = parent.add("group");
        arrowIcon.preferredSize = ARROW_ICON_SIZE;
        arrowIcon.minimumSize = ARROW_ICON_SIZE;
        arrowIcon.maximumSize = ARROW_ICON_SIZE;
        arrowIcon.helpTip = arrowName + "\n" + getLabel("tooltip.favoriteArrow");
        arrowIcon.arrowName = arrowName;
        arrowIcon.value = false;
        /* drawOptionIcon() で描くための値 / values drawOptionIcon() reads */
        arrowIcon.iconSize = ARROW_ICON_SIZE;
        arrowIcon.iconShape = ARROW_ICON_SHAPES[favoriteArrow.number];
        arrowIcon.contentInset = ARROW_ICON_INSET;
        arrowIcon.shapeScale = ARROW_ICON_SCALE;
        arrowIcon.hasLeftEdge = true;
        arrowIcon.onDraw = function () { drawOptionIcon(arrowIcon); };
        arrowIcon.addEventListener("mousedown", function () {
            if (typeof arrowIcon.onClick === "function") arrowIcon.onClick();
        });
        return arrowIcon;
    }

    /**
     * よく使う矢印のアイコンを描き直す（value を変えても自動では描き直されない）。
     * ［終点も同じ］がオンなら両端に矢印のある絵にする
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @returns {void}
     */
    function redrawFavoriteArrowIcons(arrowheadControls) {
        var showBothEnds = arrowheadControls.sameEndToggle.value;
        for (var i = 0; i < arrowheadControls.favoriteArrowIcons.length; i++) {
            arrowheadControls.favoriteArrowIcons[i].showBothEnds = showBothEnds;
            redrawStepperGroup(arrowheadControls.favoriteArrowIcons[i]);
        }
    }

    /**
     * オンになっているアイコンの選択肢を返す
     * @param {Group[]} optionIcons - addOptionIcons() で作ったアイコン
     * @returns {Object} 選択肢。どれもオフなら先頭
     */
    function getCheckedOption(optionIcons) {
        for (var i = 0; i < optionIcons.length; i++) {
            if (optionIcons[i].value) return optionIcons[i].optionDefinition;
        }
        return optionIcons[0].optionDefinition;
    }

    /**
     * 項目名・∧∨・入力欄・単位をひと組にした数値欄を追加する（ステップボタン部品の addSteppedField を使う）
     * @param {Panel|Group} parent - 追加先
     * @param {Object} fieldOptions - labelKey / labelWidth / text / min / integer / unitKey / unitInField（true で単位を欄の中に入れる）/ tooltipKey
     * @returns {EditText} 入力欄
     */
    function addNumberField(parent, fieldOptions) {
        var unitInField = !!(fieldOptions.unitInField && fieldOptions.unitKey);
        var fieldUnit = unitInField ? " " + getLabel(fieldOptions.unitKey) : undefined;
        var numberInput = addSteppedField(parent, {
            label: labelText(fieldOptions.labelKey),
            labelWidth: fieldOptions.labelWidth,
            text: unitInField ? fieldOptions.text + fieldUnit : fieldOptions.text,
            characters: unitInField ? UNIT_FIELD_CHARACTERS : FIELD_CHARACTERS,
            min: fieldOptions.min,
            integer: fieldOptions.integer,
            unit: fieldUnit,
            onStep: notifySteppedInput
        });
        numberInput.helpTip = getLabel(fieldOptions.tooltipKey);
        /* 項目名のコロンと∧∨は隙間なしで続け、単位を後ろに足す / butt the stepper against the label's colon and append the unit */
        var fieldRow = numberInput.parent.parent;
        fieldRow.spacing = STEPPER_LABEL_GAP;
        if (fieldOptions.unitKey && !unitInField) fieldRow.add("statictext", undefined, getLabel(fieldOptions.unitKey));
        return numberInput;
    }

    /**
     * 数値欄に値を書く（直接入力が不正だったときに戻す値もそろえる）
     * @param {EditText} numberInput - addNumberField() で作った入力欄
     * @param {number} value - 書く値
     * @returns {void}
     */
    function setNumberFieldValue(numberInput, value) {
        writeSteppedValue(numberInput, value, numberInput.stepperGroup.stepOptions);
    }

    /**
     * 数値欄の確定時の処理を足す（addSteppedField の値の整え直しを先に行う）
     * @param {EditText} numberInput - addNumberField() で作った入力欄
     * @param {Function} commitHandler - 整えたあとに呼ぶ処理
     * @returns {void}
     */
    function addCommitHandler(numberInput, commitHandler) {
        var normalizeInput = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeInput();
            commitHandler();
        };
    }

    /**
     * ∧∨・↑↓キーで値を変えたあと、手入力と同じ処理を呼ぶ（プログラムからの書き換えではイベントが発火しないため）
     * @param {EditText} numberInput - 値を変えた入力欄
     * @returns {void}
     */
    function notifySteppedInput(numberInput) {
        if (typeof numberInput.onChanging === "function") numberInput.onChanging();
        /* フォーカスが無いまま∧∨で変えると onChange が来ないので、確定側のプレビューもここで更新する
           The stepper can change the value without focus, so run the commit handler too */
        if (typeof numberInput.onChange === "function") numberInput.onChange();
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 線パネル（線幅・線端・角の形状）を作る
     * @param {Window|Group} parent - 追加先
     * @returns {Object} 線パネルのコントロール
     */
    function buildStrokePanel(parent) {
        var strokePanel = parent.add("panel", undefined, getLabel("panel.stroke"));
        setupPanel(strokePanel, 6);

        var strokeWidthInput = addNumberField(strokePanel, {
            labelKey: "fieldLabel.strokeWidth", text: String(getDefaultStrokeWidth()),
            min: 0, unitKey: "unit.point", unitInField: true, tooltipKey: "tooltip.strokeWidth"
        });
        var strokeWidthPopupButton = addStrokeWidthPopupButton(strokeWidthInput);
        /* 線端・角の形状は Illustrator の線パネルと同じくアイコンで選ぶ / caps and corners are picked with icons, like the Stroke panel */
        var strokeCapRow = addLabeledRow(strokePanel, "fieldLabel.strokeCap");
        var strokeCapIcons = addOptionIcons(strokeCapRow, STROKE_CAP_OPTIONS, DEFAULT_STROKE_CAP, STROKE_ICON_SIZE, STROKE_ICON_INSET);
        var cornerJoinRow = addLabeledRow(strokePanel, "fieldLabel.cornerJoin");
        var cornerJoinIcons = addOptionIcons(cornerJoinRow, CORNER_JOIN_OPTIONS, DEFAULT_CORNER_JOIN, STROKE_ICON_SIZE, STROKE_ICON_INSET);
        /* カラーはパネルの最下部 / Color sits at the bottom of the panel */
        var strokeColorRow = addLabeledRow(strokePanel, "fieldLabel.strokeColor");
        var strokeColorSwatch = addColorSwatch(strokeColorRow);
        /* 項目名の幅は実際の文字幅のいちばん広いもの（「角の形状 :」）にそろえる / match the labels to the widest actual text */
        alignLabelWidths([strokeWidthInput.fieldLabel, strokeColorRow.rowLabel, strokeCapRow.rowLabel, cornerJoinRow.rowLabel], "right");

        return {
            strokeWidthInput: strokeWidthInput,
            strokeWidthPopupButton: strokeWidthPopupButton,
            strokeColorSwatch: strokeColorSwatch,
            strokeCapRow: strokeCapRow,
            strokeCapIcons: strokeCapIcons,
            cornerJoinIcons: cornerJoinIcons
        };
    }

    /**
     * 矢印パネル（よく使う矢印・その他の矢印・倍率・オプション）を作る
     * @param {Window|Group} parent - 追加先
     * @returns {Object} 矢印パネルのコントロール
     */
    function buildArrowheadPanel(parent) {
        var arrowheadPanel = parent.add("panel", undefined, getLabel("panel.arrowhead"));
        setupPanel(arrowheadPanel, 6);
        var defaultFavoriteIndex = getDefaultFavoriteIndex();
        var arrowChoices = buildArrowChoices(arrowheadPanel, defaultFavoriteIndex);

        /* 0 以下は既定の倍率に戻されるので、下限は 1 / values of 0 or less fall back to the default, so the minimum is 1 */
        var arrowScaleGroup = arrowheadPanel.add("group");
        arrowScaleGroup.orientation = "column";
        arrowScaleGroup.alignChildren = ["fill", "top"];
        arrowScaleGroup.alignment = ["fill", "top"]; /* 縦に伸びるとオプションの上に空きができる / stretching would leave a gap above the options */
        var arrowScaleInput = addNumberField(arrowScaleGroup, {
            labelKey: "fieldLabel.arrowScale", labelWidth: ARROW_LABEL_WIDTH, text: String(FAVORITE_ARROWS[defaultFavoriteIndex].scale),
            min: 1, unitKey: "unit.percent", unitInField: true, tooltipKey: "tooltip.arrowScale"
        });

        var arrowOptions = buildArrowOptionsRow(arrowheadPanel, defaultFavoriteIndex);
        return {
            favoriteArrowIcons: arrowChoices.favoriteArrowIcons,
            otherArrowRadio: arrowChoices.otherArrowRadio,
            otherArrowList: arrowChoices.otherArrowList,
            arrowScaleInput: arrowScaleInput,
            sameEndToggle: arrowOptions.sameEndToggle,
            swapEndsToggle: arrowOptions.swapEndsToggle,
            tipAlignIcons: arrowOptions.tipAlignIcons
        };
    }

    /**
     * 矢印の選択肢（よく使う矢印のアイコンと、その他の矢印のラジオ＋ポップアップメニュー）を作る
     * @param {Panel} arrowheadPanel - 矢印パネル
     * @param {number} defaultFavoriteIndex - 最初に選んでおく FAVORITE_ARROWS の添字
     * @returns {Object} favoriteArrowIcons / otherArrowRadio / otherArrowList
     */
    function buildArrowChoices(arrowheadPanel, defaultFavoriteIndex) {
        /* よく使う矢印とその他の行は間隔0で続ける（ポップアップメニューの高さで空きが広がるため）
           Favorites and the Others row sit flush; the pop-up's height would otherwise widen the gap */
        var arrowChoiceGroup = arrowheadPanel.add("group");
        arrowChoiceGroup.orientation = "column";
        arrowChoiceGroup.alignChildren = ["fill", "top"];
        arrowChoiceGroup.alignment = ["fill", "top"]; /* 縦は伸ばさない（伸びると下に空きができる）/ do not stretch vertically */
        arrowChoiceGroup.spacing = 0;
        var favoriteArrowGroup = arrowChoiceGroup.add("group");
        favoriteArrowGroup.orientation = "column";
        favoriteArrowGroup.alignChildren = ["fill", "top"];
        favoriteArrowGroup.spacing = 6;
        favoriteArrowGroup.margins = [0, 0, 0, ARROW_ICONS_BOTTOM_MARGIN]; /* 最後の矢印とその他の行の間 / between the last icon and the Others row */

        /* よく使う矢印はアイコンで選ぶ（ラジオボタンと同じ value を持つ。排他は onClick で切り替える）
           Favorites are picked with icons that carry a radio-like value; onClick keeps them exclusive */
        /* ［なし］も含めて ARROW_ICON_COLUMNS 個ずつ横に並べる / ARROW_ICON_COLUMNS per row, [None] included */
        var favoriteArrowNames = buildFavoriteArrowNames();
        var favoriteArrowIcons = [];
        var arrowIconRow = null;
        for (var i = 0; i < favoriteArrowNames.length; i++) {
            if (i % ARROW_ICON_COLUMNS === 0) {
                arrowIconRow = favoriteArrowGroup.add("group");
                setupRow(arrowIconRow, "left", favoriteArrowGroup.spacing);
            }
            favoriteArrowIcons.push(addFavoriteArrowIcon(arrowIconRow, FAVORITE_ARROWS[i], favoriteArrowNames[i]));
        }
        favoriteArrowIcons[defaultFavoriteIndex].value = true;

        /* その他：ラジオ＋ポップアップメニュー。よく使う矢印とは別の行なので、排他は onClick で切り替える
           Others: radio + pop-up. It sits apart from the favorites, so onClick keeps them exclusive */
        var otherArrowRow = arrowChoiceGroup.add("group");
        setupRow(otherArrowRow, "left", 4);
        var otherArrowTooltip = getLabel("tooltip.otherArrow", { scale: DEFAULT_ARROW_SCALE });
        var otherArrowRadio = otherArrowRow.add("radiobutton", undefined, "");
        otherArrowRadio.helpTip = otherArrowTooltip;
        var otherArrowList = otherArrowRow.add("dropdownlist", undefined, buildOtherArrowNames());
        otherArrowList.helpTip = otherArrowTooltip;
        otherArrowList.selection = 0;
        return { favoriteArrowIcons: favoriteArrowIcons, otherArrowRadio: otherArrowRadio, otherArrowList: otherArrowList };
    }

    /**
     * 矢印のオプション（終点も同じ・入れ替え・先端位置）のアイコンを1行に並べる
     * @param {Panel} arrowheadPanel - 矢印パネル
     * @param {number} defaultFavoriteIndex - 最初に選んでおく FAVORITE_ARROWS の添字（先端位置の初期値に使う）
     * @returns {Object} sameEndToggle / swapEndsToggle / tipAlignIcons
     */
    function buildArrowOptionsRow(arrowheadPanel, defaultFavoriteIndex) {
        var arrowOptionsRow = arrowheadPanel.add("group");
        /* 縦は上寄せ。中央だとパネルが破線パネルの高さまで伸びたとき、余りの真ん中に置かれて上に空きができる
           Top-aligned; centered, it would sit in the middle of the extra height when the panel stretches */
        setupRow(arrowOptionsRow, ["left", "top"], ROW_SPACING);
        arrowOptionsRow.margins = [0, ARROW_OPTIONS_TOP_MARGIN, 0, 0]; /* 倍率との間 / space below the scale */
        return {
            sameEndToggle: addIconToggle(arrowOptionsRow, addLinkToggle, "tooltip.sameEnd", LINK_CHAIN_RATIO),
            swapEndsToggle: addIconToggle(arrowOptionsRow, addSwapToggle, "tooltip.swapEnds"),
            /* Illustrator の線パネルと同じく「終点から」「終点に」の順 / same order as Illustrator's Stroke panel */
            tipAlignIcons: addOptionIcons(arrowOptionsRow, [TIP_ALIGN_OPTIONS[1], TIP_ALIGN_OPTIONS[0]],
                FAVORITE_ARROWS[defaultFavoriteIndex].tipAlign || TIP_ALIGN_OPTIONS[0].key, TIP_ICON_SIZE)
        };
    }

    /**
     * オプションのアイコンのトグル（リンク・⇄）を追加する。切り替えたあとはアイコンの onClick を呼ぶ
     * （ラジオボタンやチェックボックスと同じつなぎ方にする）
     * @param {Group} parent - 追加先
     * @param {Function} addToggle - アイコンを作る関数（addLinkToggle / addSwapToggle）
     * @param {string} tooltipKey - helpTip の LABELS パス
     * @param {number} [drawRatio] - 絵の大きさの比率（addLinkToggle の鎖の比率。省略時は 1）
     * @returns {Group} アイコン（.value でオンかを読む）
     */
    function addIconToggle(parent, addToggle, tooltipKey, drawRatio) {
        var iconToggle = addToggle(parent, false, function () {
            if (typeof iconToggle.onClick === "function") iconToggle.onClick();
        }, OPTION_TOGGLE_SIZE, drawRatio);
        iconToggle.helpTip = getLabel(tooltipKey);
        return iconToggle;
    }

    /**
     * 始点と終点の入れ替えアイコン（⇄）を追加する。見た目と操作はリンクアイコンにそろえる
     * （オンのときは押し込んだボタンのように地と枠を描く。配色・描き直しはリンクアイコンの部品を使う）
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - [幅, 高さ]（省略時は LINK_ICON_SIZE）
     * @returns {Group} アイコン（.value でオンかを読む）
     */
    function addSwapToggle(parent, initialValue, onToggle, iconSize) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var swapToggle = parent.add("group");
        swapToggle.preferredSize = toggleSize;
        swapToggle.minimumSize = toggleSize;
        swapToggle.maximumSize = toggleSize;
        swapToggle.value = initialValue;

        swapToggle.onDraw = function () {
            var iconGraphics = swapToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            var isDimmed = !isLinkToggleEnabledInTree(swapToggle);
            if (swapToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawSwapIcon(iconGraphics, iconWidth, iconHeight, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
        };

        swapToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(swapToggle)) return;
            swapToggle.value = !swapToggle.value;
            redrawLinkToggle(swapToggle);
            if (onToggle) onToggle();
        });
        return swapToggle;
    }

    /**
     * ⇄ を描く（22px 基準の絵を大きさに合わせて拡大し、中央に置く。右向きの矢印を上、左向きの矢印を下）。
     * ScriptUI は多角形を塗れないため、矢じりは細い長方形を並べて三角形に近づける
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSwapIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        /** 22px 基準の x を描画先へ / map a 22 px based x */
        function mapX(x) { return offsetX + x * iconScale; }
        /** 22px 基準の y を描画先へ / map a 22 px based y */
        function mapY(y) { return offsetY + y * iconScale; }
        iconGraphics.newPath();
        iconGraphics.rectPath(mapX(5.5), mapY(7.2), 6 * iconScale, 2 * iconScale);   /* 上の矢印の軸 / upper shaft */
        addArrowHeadPath(iconGraphics, mapX(11.5), mapX(17), mapY(8.2), 2.8 * iconScale);
        iconGraphics.rectPath(mapX(9.3), mapY(13), 6 * iconScale, 2 * iconScale);    /* 下の矢印の軸 / lower shaft */
        addArrowHeadPath(iconGraphics, mapX(9.3), mapX(3.8), mapY(14), 2.8 * iconScale);
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 円か円の一部（全体・左半分・右半分・左上の四分円）を細い長方形の並びでパスに足す（ScriptUI は円弧を塗れないため）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} centerX - 中心の x
     * @param {number} centerY - 中心の y
     * @param {number} radius - 半径
     * @param {string} discSide - "full"（全体）/ "left"（左半分）/ "right"（右半分）/ "topLeft"（左上の四分円）
     * @returns {void}
     */
    function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
        var sliceCount = Math.max(8, Math.ceil(radius * 2));
        var sliceHeight = radius / sliceCount;
        var bottomSlices = (discSide === "topLeft") ? 0 : sliceCount;
        for (var k = -sliceCount; k < bottomSlices; k++) {
            /* 帯の中央の高さで円の幅を測る / measure the half width at the slice's middle */
            var sliceCenterY = (k + 0.5) * sliceHeight;
            var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
            var sliceLeft = (discSide === "right") ? centerX : centerX - halfWidth;
            var sliceWidth = (discSide === "full") ? halfWidth * 2 : halfWidth;
            iconGraphics.rectPath(sliceLeft, centerY + k * sliceHeight, sliceWidth, sliceHeight);
        }
    }

    /**
     * 矢じり（三角形）を細い長方形の並びでパスに足す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} baseX - 矢じりの付け根の x
     * @param {number} tipX - 先端の x
     * @param {number} centerY - 中心の y
     * @param {number} halfHeight - 付け根の高さの半分
     * @returns {void}
     */
    function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
        var sliceCount = 12;
        var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
        var direction = (tipX > baseX) ? 1 : -1;
        for (var k = 0; k < sliceCount; k++) {
            var sliceStart = baseX + direction * sliceWidth * k;
            var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
            iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
        }
    }

    /**
     * 破線パネル（破線の種類・破線の計算・計算方法・両端を調整）を作る
     * @param {Window|Group} parent - 追加先
     * @returns {Object} 破線パネルのコントロール
     */
    function buildDashPanel(parent) {
        var dashPanel = parent.add("panel", undefined, getLabel("panel.dash"));
        setupPanel(dashPanel, 6);
        dashPanel.alignment = ["fill", "top"]; /* 左の列の高さまで伸ばさず上に揃える / top-aligned, not stretched to the left column */

        /* なし・破線・ドット点線は縦に並べる。ほかのラジオと混ざらないようグループに入れる（同じ親の中だけ排他）
           None / Dashed / Dotted stacked in their own group (radios are exclusive only within one parent) */
        var dashStyleGroup = dashPanel.add("group");
        dashStyleGroup.orientation = "column";
        dashStyleGroup.alignChildren = ["left", "top"];
        dashStyleGroup.spacing = 6;
        var noDashRadio = addOptionRadio(dashStyleGroup, getLabel("radio.noDash"), getLabel("tooltip.noDash"));
        var dashedRadio = addOptionRadio(dashStyleGroup, getLabel("radio.dashed"), getLabel("tooltip.dashed"));
        var dottedRadio = addOptionRadio(dashStyleGroup, getLabel("radio.dotted"), getLabel("tooltip.dotted"));
        noDashRadio.value = true;

        /* 破線の計算（DashGapCalculator から移植）/ Dash calculation (ported from DashGapCalculator) */
        /* 破線の計算は枠で囲まず、上に区切り線を引いて分ける / the dash calculation sits below a separator, without a frame */
        addSeparator(dashPanel);
        var dashCalcGroup = dashPanel.add("group");
        dashCalcGroup.orientation = "column";
        dashCalcGroup.alignChildren = ["fill", "top"];
        dashCalcGroup.alignment = ["fill", "top"];
        dashCalcGroup.spacing = 6;
        dashCalcGroup.margins = [0, SEPARATOR_BOTTOM_MARGIN, 0, 0]; /* 区切り線の下 / below the separator */
        var segmentsInput = addNumberField(dashCalcGroup, {
            labelKey: "fieldLabel.segments", text: "1",
            min: 1, integer: true, tooltipKey: "tooltip.segments"
        });
        var gapInput = addNumberField(dashCalcGroup, {
            labelKey: "fieldLabel.gap", text: "0",
            min: 0, unitKey: "unit.point", unitInField: true, tooltipKey: "tooltip.gap"
        });
        var dashLengthInput = addNumberField(dashCalcGroup, {
            labelKey: "fieldLabel.dash", text: "0",
            min: 0, unitKey: "unit.point", unitInField: true, tooltipKey: "tooltip.dash"
        });
        /* 項目名の幅は実際の文字幅のいちばん広いものにそろえる / match the label widths to the widest actual text */
        alignLabelWidths([segmentsInput.fieldLabel, gapInput.fieldLabel, dashLengthInput.fieldLabel], "right");

        /* 計算方法と両端を調整は破線の計算の中に置く / Calculation and Adjust ends sit inside Dash Calculation */
        var calcMethodPanel = addSubPanel(dashCalcGroup, "panel.calcMethod");
        var gapToDashRadio = addOptionRadio(calcMethodPanel, getLabel("radio.gapToDash"), getLabel("tooltip.gapToDash"));
        var dashToGapRadio = addOptionRadio(calcMethodPanel, getLabel("radio.dashToGap"), getLabel("tooltip.dashToGap"));
        gapToDashRadio.value = true;

        /* 角に合わせる（両端を調整）のアイコンは左右中央に置く / the adjust-ends icons sit centered */
        var adjustDashEndsIcons = addOptionIcons(dashCalcGroup, ADJUST_DASH_OPTIONS, "adjustEnds", ADJUST_ICON_SIZE, ADJUST_ICON_INSET);
        adjustDashEndsIcons[0].parent.alignment = ["center", "top"];

        return {
            noDashRadio: noDashRadio,
            dashedRadio: dashedRadio,
            dottedRadio: dottedRadio,
            dashCalcGroup: dashCalcGroup,
            segmentsInput: segmentsInput,
            gapInput: gapInput,
            dashLengthInput: dashLengthInput,
            calcMethodPanel: calcMethodPanel,
            gapToDashRadio: gapToDashRadio,
            dashToGapRadio: dashToGapRadio,
            adjustDashEndsIcons: adjustDashEndsIcons
        };
    }

    /**
     * 最上部のプリセットの行（ドロップダウン・保存・削除）を作る
     * @param {Window} parent - 追加先
     * @returns {Object} プリセットの行のコントロール
     */
    function buildPresetRow(parent) {
        var presetRow = parent.add("group");
        setupRow(presetRow, "center", ROW_SPACING); /* ダイアログボックスの左右中央に置く / centered in the dialog */
        presetRow.add("statictext", undefined, labelText("fieldLabel.preset"));
        var presetDropdown = presetRow.add("dropdownlist", undefined, []);
        presetDropdown.helpTip = getLabel("tooltip.preset");
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        /* 保存アイコンは塗りの面が大きく濃く見えるので、色を薄めて描く / the save icon is mostly solid and looks heavy, so draw it lighter */
        var presetSaveIcon = addIconButton(presetRow, PRESET_ICON_SIZE, function (iconGraphics, iconWidth, iconHeight, iconColor) {
            drawSaveIcon(iconGraphics, iconWidth, iconHeight, [iconColor[0], iconColor[1], iconColor[2], iconColor[3] * SAVE_ICON_OPACITY]);
        });
        presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
        var presetDeleteIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel("tooltip.presetDelete");
        return {
            presetDropdown: presetDropdown,
            presetSaveIcon: presetSaveIcon,
            presetDeleteIcon: presetDeleteIcon
        };
    }

    /**
     * 設定用ダイアログのコントロールをすべて作る（イベントは設定しない）
     * @returns {Object} ダイアログとパネルごとのコントロール
     */
    function buildSettingsDialog() {
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        var presetControls = buildPresetRow(settingsDialog);

        /* 2 カラム：左に線と矢印、右に破線 / Two columns: stroke and arrowheads on the left, dashes on the right */
        var settingsColumns = settingsDialog.add("group");
        settingsColumns.orientation = "row";
        settingsColumns.alignChildren = ["fill", "fill"];
        settingsColumns.spacing = COLUMN_SPACING;
        var leftColumn = settingsColumns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        var strokeControls = buildStrokePanel(leftColumn);
        var arrowheadControls = buildArrowheadPanel(leftColumn);
        var dashControls = buildDashPanel(settingsColumns);

        /* ボタン（左：「線」パネルを開く／右：キャンセル・OK）。プレビューは常にオン
           Buttons (left: Open Stroke Panel, right: cancel and OK); the preview is always on */
        var buttonRow = addButtonRow(settingsDialog);
        var openStrokePanelButton = buttonRow.leftGroup.add("button", undefined, getLabel("button.openStrokePanel"));
        openStrokePanelButton.helpTip = getLabel("tooltip.openStrokePanel");
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            settingsDialog: settingsDialog,
            preset: presetControls,
            stroke: strokeControls,
            openStrokePanelButton: openStrokePanelButton,
            arrowhead: arrowheadControls,
            dash: dashControls,
        };
    }

    // =========================================
    // ダイアログの入力値 / Dialog values
    // =========================================

    /**
     * 選択中の矢印名を返す（よく使う矢印はアイコンの .arrowName、その他はメニューの表示名）
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @returns {string} 矢印名
     */
    function getSelectedArrowName(arrowheadControls) {
        if (arrowheadControls.otherArrowRadio.value) return arrowheadControls.otherArrowList.selection.text;
        var favoriteArrowIcons = arrowheadControls.favoriteArrowIcons;
        for (var i = 0; i < favoriteArrowIcons.length; i++) {
            if (favoriteArrowIcons[i].value) return favoriteArrowIcons[i].arrowName;
        }
        return favoriteArrowIcons[getDefaultFavoriteIndex()].arrowName;
    }

    /**
     * 破線の計算の入力値を読む
     * @param {Object} dashControls - buildDashPanel() の戻り値
     * @returns {Object|null} calcDashArray() に渡す値。破線がオフなら null
     */
    function readDashCalc(dashControls) {
        if (dashControls.noDashRadio.value) return null;
        return {
            style: dashControls.dottedRadio.value ? "dotted" : "dashed",
            mode: dashControls.dashToGapRadio.value ? "dashToGap" : "gapToDash",
            segments: parseInt(dashControls.segmentsInput.text, 10),
            gapPt: parseFloat(dashControls.gapInput.text),
            dashPt: parseFloat(dashControls.dashLengthInput.text),
            adjustEnds: (getCheckedOption(dashControls.adjustDashEndsIcons).key === "adjustEnds")
        };
    }

    /**
     * ダイアログの入力値を、アクションと破線に渡す設定にまとめる
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @returns {Object|null} 線の設定。線幅が不正なら null
     */
    function readDialogSettings(dialogControls) {
        var arrowheadControls = dialogControls.arrowhead;
        var strokeWidth = parseFloat(dialogControls.stroke.strokeWidthInput.text);
        if (isNaN(strokeWidth) || strokeWidth < 0) return null;

        var arrowScale = parseFloat(arrowheadControls.arrowScaleInput.text);
        if (isNaN(arrowScale) || arrowScale <= 0) arrowScale = DEFAULT_ARROW_SCALE;

        /* 始点に付けるのが基本。終点も同じなら両端、入れ替えなら終点だけ
           The start gets the arrowhead; same-at-end puts it on both ends, swap moves it to the end */
        var arrowName = getSelectedArrowName(arrowheadControls);
        var noArrowName = getLabel("actionName.noArrowhead");
        var isSwapped = !arrowheadControls.sameEndToggle.value && arrowheadControls.swapEndsToggle.value;
        var hasEndArrow = arrowheadControls.sameEndToggle.value || isSwapped;

        var dashCalc = readDashCalc(dialogControls.dash);
        /* ドット点線は長さ0の線分なので、丸型線端でないと見えない / Dots are zero-length dashes, visible only with round caps */
        var isDotted = (dashCalc !== null && dashCalc.style === "dotted");

        return {
            strokeWidth: strokeWidth,
            startArrow: isSwapped ? noArrowName : arrowName,
            startScale: isSwapped ? DEFAULT_ARROW_SCALE : arrowScale,
            endArrow: hasEndArrow ? arrowName : noArrowName,
            endScale: hasEndArrow ? arrowScale : DEFAULT_ARROW_SCALE,
            tipAlign: getCheckedOption(arrowheadControls.tipAlignIcons),
            strokeCap: isDotted ? findOptionByKey(STROKE_CAP_OPTIONS, "round") : getCheckedOption(dialogControls.stroke.strokeCapIcons),
            cornerJoin: getCheckedOption(dialogControls.stroke.cornerJoinIcons),
            dashCalc: dashCalc,
            strokeColor: dialogControls.stroke.strokeColorSwatch.isChanged ? dialogControls.stroke.strokeColorSwatch.strokeColor : null
        };
    }

    /**
     * 先頭のパスで破線を計算し、求める側の欄に結果を書く
     * @param {Object} dashControls - buildDashPanel() の戻り値
     * @param {Object|null} pathMetrics - getFirstPathMetrics() の戻り値
     * @returns {void}
     */
    function refreshDashCalcResult(dashControls, pathMetrics) {
        var dashCalc = readDashCalc(dashControls);
        if (!dashCalc || !pathMetrics) return;
        var dashArray = calcDashArray(dashCalc, pathMetrics.length, pathMetrics.closed);
        if (!dashArray) return;
        var isDotted = (dashCalc.style === "dotted");
        if (isDotted || dashCalc.mode === "dashToGap") setNumberFieldValue(dashControls.gapInput, dashArray[1]);
        if (isDotted || dashCalc.mode === "gapToDash") setNumberFieldValue(dashControls.dashLengthInput, dashArray[0]);
    }

    /**
     * オンにした破線の種類に合わせて、線幅比のパターンから分割数・間隔・線分の初期値を入れる
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @param {Object|null} pathMetrics - getFirstPathMetrics() の戻り値
     * @returns {void}
     */
    function fillDashCalcDefaults(dialogControls, pathMetrics) {
        var dashControls = dialogControls.dash;
        var dashCalc = readDashCalc(dashControls);
        if (!dashCalc) return;
        var dashPattern = DASH_PATTERNS[dashCalc.style];
        var strokeWidth = parseFloat(dialogControls.stroke.strokeWidthInput.text);
        if (isNaN(strokeWidth) || strokeWidth <= 0) strokeWidth = getDefaultStrokeWidth();
        if (pathMetrics) {
            setNumberFieldValue(dashControls.segmentsInput, estimateSegments(dashCalc.style, strokeWidth, pathMetrics, dashCalc.adjustEnds));
        }
        setNumberFieldValue(dashControls.gapInput, dashPattern.gap * strokeWidth);
        setNumberFieldValue(dashControls.dashLengthInput, dashPattern.dash * strokeWidth);
    }

    /**
     * 矢印のオプション（終点も同じ・入れ替え・先端位置）の有効／無効をそろえる。
     * 矢印が［なし］ならすべてディム、［終点も同じ］がオンなら入れ替えをディム
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @returns {void}
     */
    function syncArrowOptionsEnabled(arrowheadControls) {
        /* 矢印を選び直したあとに呼ばれるので、よく使う矢印のアイコンもここで描き直す / also repaint the favorite icons */
        redrawFavoriteArrowIcons(arrowheadControls);
        var hasArrow = (getSelectedArrowNumber(arrowheadControls) !== 0);
        setLinkToggleEnabled(arrowheadControls.sameEndToggle, hasArrow);
        setLinkToggleEnabled(arrowheadControls.swapEndsToggle, hasArrow && !arrowheadControls.sameEndToggle.value);
        setOptionIconsEnabled(arrowheadControls.tipAlignIcons, hasArrow);
    }

    /**
     * 破線の有無・種類・計算方法に合わせて有効／無効を切り替える。
     * 入力する側の欄だけ有効にし、ドット点線は分割数だけ。ドット点線の間は線端を丸型に固定するのでディム
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @returns {void}
     */
    function syncDashEnabled(dialogControls) {
        var dashControls = dialogControls.dash;
        var isDotted = dashControls.dottedRadio.value;
        var hasDash = !dashControls.noDashRadio.value;
        var isDashed = hasDash && !isDotted;
        dialogControls.stroke.strokeCapRow.enabled = !isDotted;
        setOptionIconsEnabled(dialogControls.stroke.strokeCapIcons, !isDotted); /* 自作描画は描き直しが要る / custom drawing needs a repaint */
        dashControls.dashCalcGroup.enabled = hasDash;
        dashControls.calcMethodPanel.enabled = isDashed;
        setOptionIconsEnabled(dashControls.adjustDashEndsIcons, hasDash);
        setSteppedFieldEnabled(dashControls.segmentsInput, hasDash);
        setSteppedFieldEnabled(dashControls.gapInput, isDashed && dashControls.gapToDashRadio.value);
        setSteppedFieldEnabled(dashControls.dashLengthInput, isDashed && dashControls.dashToGapRadio.value);
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /**
     * 保存したプリセットを返す
     * @returns {Object} プリセット名をキーにした設定の集まり（書き換えても保存されない）
     */
    function loadPresetMap() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットの名前を、名前順で返す
     * @returns {string[]} プリセット名
     */
    function getPresetNames() {
        var presetMap = loadPresetMap();
        var presetNames = [];
        for (var presetName in presetMap) {
            if (presetMap.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * 選択中の矢印の番号を返す（[なし] は 0）
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @returns {number} 矢印の番号
     */
    function getSelectedArrowNumber(arrowheadControls) {
        if (arrowheadControls.otherArrowRadio.value) {
            return parseInt(arrowheadControls.otherArrowList.selection.text.replace(/[^0-9]/g, ""), 10);
        }
        var favoriteArrowIcons = arrowheadControls.favoriteArrowIcons;
        for (var i = 0; i < favoriteArrowIcons.length; i++) {
            if (favoriteArrowIcons[i].value) return FAVORITE_ARROWS[i].number;
        }
        return FAVORITE_ARROWS[getDefaultFavoriteIndex()].number;
    }

    /**
     * 番号で矢印を選ぶ。よく使う矢印に無ければポップアップメニューから選ぶ
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @param {number} arrowNumber - 矢印の番号（[なし] は 0）
     * @returns {void}
     */
    function selectArrowByNumber(arrowheadControls, arrowNumber) {
        var favoriteArrowIcons = arrowheadControls.favoriteArrowIcons;
        var isFavorite = false;
        for (var i = 0; i < favoriteArrowIcons.length; i++) {
            favoriteArrowIcons[i].value = (FAVORITE_ARROWS[i].number === arrowNumber);
            if (favoriteArrowIcons[i].value) isFavorite = true;
        }
        arrowheadControls.otherArrowRadio.value = !isFavorite;
        if (isFavorite) return;
        var otherArrowName = getLabel("actionName.arrowheadPrefix") + arrowNumber;
        var otherArrowItems = arrowheadControls.otherArrowList.items;
        for (var j = 0; j < otherArrowItems.length; j++) {
            if (otherArrowItems[j].text === otherArrowName) {
                arrowheadControls.otherArrowList.selection = j;
                return;
            }
        }
    }

    /**
     * key が一致するアイコンの位置を返す
     * @param {Group[]} optionIcons - addOptionIcons() で作ったアイコン
     * @param {string} optionKey - 選択肢の key
     * @returns {number} 添字。無ければ -1
     */
    function findOptionIndexByKey(optionIcons, optionKey) {
        for (var i = 0; i < optionIcons.length; i++) {
            if (optionIcons[i].optionDefinition.key === optionKey) return i;
        }
        return -1;
    }

    /**
     * キーが一致するアイコンをオンにして描き直す（value を変えても自動では描き直されない）
     * @param {Group[]} optionIcons - addOptionIcons() で作ったアイコン
     * @param {string} optionKey - 選択肢の key
     * @returns {void}
     */
    function checkOptionByKey(optionIcons, optionKey) {
        for (var i = 0; i < optionIcons.length; i++) {
            optionIcons[i].value = (optionIcons[i].optionDefinition.key === optionKey);
            redrawStepperGroup(optionIcons[i]);
        }
    }

    /**
     * ダイアログの値をプリセットの形にまとめる（矢印は言語に依らない番号で持つ）
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @returns {Object} プリセット
     */
    function collectPresetData(dialogControls) {
        var strokeControls = dialogControls.stroke;
        var arrowheadControls = dialogControls.arrowhead;
        var dashControls = dialogControls.dash;
        return {
            strokeWidth: parseFloat(strokeControls.strokeWidthInput.text),
            strokeCap: getCheckedOption(strokeControls.strokeCapIcons).key,
            cornerJoin: getCheckedOption(strokeControls.cornerJoinIcons).key,
            arrowNumber: getSelectedArrowNumber(arrowheadControls),
            arrowScale: parseFloat(arrowheadControls.arrowScaleInput.text),
            sameEnd: arrowheadControls.sameEndToggle.value,
            swapEnds: arrowheadControls.swapEndsToggle.value,
            tipAlign: getCheckedOption(arrowheadControls.tipAlignIcons).key,
            dashStyle: dashControls.dottedRadio.value ? "dotted" : (dashControls.dashedRadio.value ? "dashed" : "none"),
            dashMode: dashControls.dashToGapRadio.value ? "dashToGap" : "gapToDash",
            segments: parseInt(dashControls.segmentsInput.text, 10),
            gap: parseFloat(dashControls.gapInput.text),
            dash: parseFloat(dashControls.dashLengthInput.text),
            adjustDashEnds: (getCheckedOption(dashControls.adjustDashEndsIcons).key === "adjustEnds")
        };
    }

    /**
     * プリセットをダイアログに書き込む（ディム表示とプレビューは呼び出し側でそろえる）
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetData(dialogControls, presetData) {
        var strokeControls = dialogControls.stroke;
        var arrowheadControls = dialogControls.arrowhead;
        var dashControls = dialogControls.dash;

        if (!isNaN(presetData.strokeWidth)) setNumberFieldValue(strokeControls.strokeWidthInput, presetData.strokeWidth);
        checkOptionByKey(strokeControls.strokeCapIcons, presetData.strokeCap);
        checkOptionByKey(strokeControls.cornerJoinIcons, presetData.cornerJoin);

        /* メニューで選ぶと倍率が 100% に戻るので、倍率は矢印のあとに書く / Picking from the menu resets the scale, so write it afterwards */
        if (typeof presetData.arrowNumber === "number") selectArrowByNumber(arrowheadControls, presetData.arrowNumber);
        if (!isNaN(presetData.arrowScale)) setNumberFieldValue(arrowheadControls.arrowScaleInput, presetData.arrowScale);
        setLinkToggleValue(arrowheadControls.sameEndToggle, !!presetData.sameEnd);
        setLinkToggleValue(arrowheadControls.swapEndsToggle, !!presetData.swapEnds);
        syncArrowOptionsEnabled(arrowheadControls);
        checkOptionByKey(arrowheadControls.tipAlignIcons, presetData.tipAlign);

        dashControls.noDashRadio.value = (presetData.dashStyle !== "dashed" && presetData.dashStyle !== "dotted");
        dashControls.dashedRadio.value = (presetData.dashStyle === "dashed");
        dashControls.dottedRadio.value = (presetData.dashStyle === "dotted");
        dashControls.dashToGapRadio.value = (presetData.dashMode === "dashToGap");
        dashControls.gapToDashRadio.value = !dashControls.dashToGapRadio.value;
        if (!isNaN(presetData.segments)) setNumberFieldValue(dashControls.segmentsInput, presetData.segments);
        if (!isNaN(presetData.gap)) setNumberFieldValue(dashControls.gapInput, presetData.gap);
        if (!isNaN(presetData.dash)) setNumberFieldValue(dashControls.dashLengthInput, presetData.dash);
        checkOptionByKey(dashControls.adjustDashEndsIcons, presetData.adjustDashEnds ? "adjustEnds" : "keepLength");
    }

    /**
     * プリセットのドロップダウンを作り直し、指定の名前を選ぶ
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @param {string} [selectedName] - 選んでおくプリセット名（省略時は「---」）
     * @returns {void}
     */
    function fillPresetDropdown(presetControls, selectedName) {
        var presetDropdown = presetControls.presetDropdown;
        var presetNames = getPresetNames();
        var selectedIndex = 0;
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel("dropdown.presetPlaceholder"));
        for (var i = 0; i < presetNames.length; i++) {
            presetDropdown.add("item", presetNames[i]);
            if (presetNames[i] === selectedName) selectedIndex = i + 1;
        }
        presetDropdown.selection = selectedIndex;
        setIconButtonEnabled(presetControls.presetDeleteIcon, selectedIndex > 0);
    }

    /**
     * 選んでいるプリセット名を返す
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @returns {string|null} プリセット名。「---」なら null
     */
    function getSelectedPresetName(presetControls) {
        var presetSelection = presetControls.presetDropdown.selection;
        return (presetSelection && presetSelection.index > 0) ? presetSelection.text : null;
    }

    /**
     * プリセット名を尋ねる
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel("dialog.presetSave"));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, "left", ROW_SPACING);
        nameRow.add("statictext", undefined, labelText("fieldLabel.presetName"));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel("tooltip.presetName");
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(nameDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;

        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        if (!presetName || presetName === getLabel("dropdown.presetPlaceholder")) return null;
        return presetName;
    }

    /**
     * 今の設定を名前を付けて保存する（同じ名前は確認してから上書き）
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function saveCurrentPreset(dialogControls) {
        var presetName = showPresetNameDialog(getSelectedPresetName(dialogControls.preset) || "");
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        if (presetMap.hasOwnProperty(presetName) && !confirm(getLabel("confirm.presetOverwrite", { name: presetName }))) return null;
        presetMap[presetName] = collectPresetData(dialogControls);
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return null;
        }
        return presetName;
    }

    /**
     * 選んでいるプリセットを確認してから削除する
     * @param {Object} presetControls - buildPresetRow() の戻り値
     * @returns {boolean} 削除したら true
     */
    function deleteSelectedPreset(presetControls) {
        var presetName = getSelectedPresetName(presetControls);
        if (!presetName || !confirm(getLabel("confirm.presetDelete", { name: presetName }))) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return false;
        }
        return true;
    }

    // =========================================
    // ダイアログの表示 / Dialog display
    // =========================================

    /**
     * 設定用ダイアログを表示し、イベントをつなぐ
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {Object|null} 線の設定。キャンセル時は null
     */
    function showSettingsDialog(targetDocument) {
        var dialogControls = buildSettingsDialog();
        var selectedPaths = collectStrokePaths(targetDocument.selection);
        /* 線幅欄は選択しているパスの線幅から始める（線が無ければ単位に合わせた初期値のまま）
           Start the weight from the selected path's stroke; without one, keep the unit-based default */
        var selectedStrokeWidth = getSelectedStrokeWidth(selectedPaths);
        if (selectedStrokeWidth !== null) setNumberFieldValue(dialogControls.stroke.strokeWidthInput, selectedStrokeWidth);
        /* 色見本は選択しているパスの色から始める（線が無ければ塗りの色）/ the swatch starts from the selected path's color */
        var selectedStrokeColor = getSelectedStrokeColor(selectedPaths);
        if (selectedStrokeColor) dialogControls.stroke.strokeColorSwatch.strokeColor = selectedStrokeColor;
        var dialogSession = {
            controls: dialogControls,
            previewController: createPreviewController(),
            firstPathMetrics: getFirstPathMetrics(selectedPaths),
            isApplyingPreset: false, /* プリセットの書き込み中はプレビューを止める / suspend the preview while a preset is written */
            strokeWidthScale: getArrowStrokeWidthScale(DEFAULT_ARROW_NUMBER) /* 今の矢印が線幅に掛けている倍数 / weight multiplier of the current arrowhead */
        };
        bindStrokeEvents(dialogSession);
        bindArrowheadEvents(dialogSession);
        bindDashEvents(dialogSession);
        bindPresetEvents(dialogSession);
        syncDashEnabled(dialogControls);
        syncArrowOptionsEnabled(dialogControls.arrowhead);

        /* プレビューは常にオン。開く前に今の設定で反映しておく / The preview is always on; show the current settings before opening */
        updatePreview(dialogSession);

        /* 「線」パネルを開くは、キャンセルと同じく閉じてから開く（モーダルの間はパネルを操作できない）
           Open Stroke Panel closes like Cancel, then opens the panel (it cannot be used while the dialog is modal) */
        var isStrokePanelRequested = false;
        dialogControls.openStrokePanelButton.onClick = function () {
            isStrokePanelRequested = true;
            dialogControls.settingsDialog.close(2);
        };

        prepareDialogWindow(dialogControls.settingsDialog, SCRIPT_NAME);
        var isAccepted = (dialogControls.settingsDialog.show() === 1);

        /* プレビューを必ず取り消してから本適用に進む / Always revert the preview before applying */
        dialogSession.previewController.reset();
        if (isStrokePanelRequested) app.executeMenuCommand("Adobe Stroke Palette");
        if (!isAccepted) return null;

        var strokeSettings = readDialogSettings(dialogControls);
        if (!strokeSettings) alert(getLabel("alert.invalidWidth"));
        return strokeSettings;
    }

    // =========================================
    // ダイアログのイベント / Dialog events
    // =========================================
    // dialogSession は { controls, previewController, firstPathMetrics, isApplyingPreset }
    // dialogSession holds { controls, previewController, firstPathMetrics, isApplyingPreset }

    /**
     * プレビューを現在の入力値で更新する（矢印を含むためアクションを実行）。プリセットの書き込み中は何もしない
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function updatePreview(dialogSession) {
        if (dialogSession.isApplyingPreset) return;
        var strokeSettings = readDialogSettings(dialogSession.controls);
        if (strokeSettings) dialogSession.previewController.previewSettings(strokeSettings);
    }

    /**
     * 破線の計算を更新してプレビューする
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function updateDashCalc(dialogSession) {
        refreshDashCalcResult(dialogSession.controls.dash, dialogSession.firstPathMetrics);
        updatePreview(dialogSession);
    }

    /**
     * 矢印を1つだけ選び、倍率を入れてプレビューする。よく使う矢印なら、その矢印の先端位置・線幅の倍数・線端・角の形状も入れる
     * （線端・角の形状の指定が無い矢印は、既定の線端なし・マイター結合に戻す）
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @param {Object} targetChoice - 選ぶよく使う矢印のアイコン、またはその他のラジオボタン
     * @param {Object|null} favoriteArrow - FAVORITE_ARROWS の要素。その他の矢印なら null（倍率は DEFAULT_ARROW_SCALE）
     * @returns {void}
     */
    function selectArrowChoice(dialogSession, targetChoice, favoriteArrow) {
        var dialogControls = dialogSession.controls;
        var arrowheadControls = dialogControls.arrowhead;
        var arrowSettings = favoriteArrow || {};
        /* プリセットの書き込み中は線幅・線端もプリセットの値なので触らない / a preset brings its own weight, cap and corner */
        if (!dialogSession.isApplyingPreset) {
            changeArrowStrokeWidthScale(dialogSession, arrowSettings.strokeWidthScale || 1);
            /* 指定の無い矢印は線端なし・マイター結合に戻す / arrowheads without their own cap and corner go back to the defaults */
            checkOptionByKey(dialogControls.stroke.strokeCapIcons, arrowSettings.strokeCap || DEFAULT_STROKE_CAP);
            checkOptionByKey(dialogControls.stroke.cornerJoinIcons, arrowSettings.cornerJoin || DEFAULT_CORNER_JOIN);
        }
        var favoriteArrowIcons = arrowheadControls.favoriteArrowIcons;
        for (var i = 0; i < favoriteArrowIcons.length; i++) favoriteArrowIcons[i].value = (favoriteArrowIcons[i] === targetChoice);
        arrowheadControls.otherArrowRadio.value = (arrowheadControls.otherArrowRadio === targetChoice);
        setNumberFieldValue(arrowheadControls.arrowScaleInput, favoriteArrow ? favoriteArrow.scale : DEFAULT_ARROW_SCALE);
        if (arrowSettings.tipAlign) checkOptionByKey(arrowheadControls.tipAlignIcons, arrowSettings.tipAlign);
        syncArrowOptionsEnabled(arrowheadControls);
        updatePreview(dialogSession);
    }

    /**
     * 矢印の番号から、その矢印が線幅に掛ける倍数を返す
     * @param {number} arrowNumber - 矢印の番号（[なし] は 0）
     * @returns {number} 倍数。よく使う矢印に無いか指定が無ければ 1
     */
    function getArrowStrokeWidthScale(arrowNumber) {
        var favoriteIndex = findFavoriteArrowIndex(arrowNumber);
        return (favoriteIndex >= 0 && FAVORITE_ARROWS[favoriteIndex].strokeWidthScale) || 1;
    }

    /**
     * 線幅欄の値から前の矢印の倍数を外し、新しい倍数を掛ける（同じ倍数なら何もしない）
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @param {number} strokeWidthScale - 新しい矢印の倍数
     * @returns {void}
     */
    function changeArrowStrokeWidthScale(dialogSession, strokeWidthScale) {
        if (strokeWidthScale === dialogSession.strokeWidthScale) return;
        var strokeWidthInput = dialogSession.controls.stroke.strokeWidthInput;
        var strokeWidth = parseFloat(strokeWidthInput.text);
        if (!isNaN(strokeWidth)) setNumberFieldValue(strokeWidthInput, strokeWidth / dialogSession.strokeWidthScale * strokeWidthScale);
        dialogSession.strokeWidthScale = strokeWidthScale;
    }

    /**
     * 線パネルのイベントをつなぐ。線幅は入力中も DOM で即時プレビューし、確定時にアクションで貼り直す
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function bindStrokeEvents(dialogSession) {
        var strokeControls = dialogSession.controls.stroke;
        strokeControls.strokeWidthInput.onChanging = function () {
            dialogSession.previewController.previewStrokeWidth(parseFloat(strokeControls.strokeWidthInput.text));
        };
        addCommitHandler(strokeControls.strokeWidthInput, function () { updatePreview(dialogSession); });
        strokeControls.strokeColorSwatch.onColorChange = function () { updatePreview(dialogSession); };
        /* ▼で選んだ値は、手入力と同じく確定の処理を通す / a picked weight goes through the same commit as typing */
        strokeControls.strokeWidthPopupButton.onPick = function (pickedWidth) {
            setNumberFieldValue(strokeControls.strokeWidthInput, pickedWidth);
            strokeControls.strokeWidthInput.onChange();
        };
        var strokeOptionIcons = strokeControls.strokeCapIcons.concat(strokeControls.cornerJoinIcons);
        for (var i = 0; i < strokeOptionIcons.length; i++) {
            strokeOptionIcons[i].onClick = function () { updatePreview(dialogSession); };
        }
        /* 丸型線端を option＋クリックすると、角の形状もラウンドにする / Option-clicking Round Cap also sets Round Join */
        var roundCapIcon = strokeControls.strokeCapIcons[findOptionIndexByKey(strokeControls.strokeCapIcons, "round")];
        roundCapIcon.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) checkOptionByKey(strokeControls.cornerJoinIcons, "round");
            updatePreview(dialogSession);
        };
    }

    /**
     * 矢印パネルのイベントをつなぐ。矢印を選んだらその矢印の倍率・先端位置を入れる
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function bindArrowheadEvents(dialogSession) {
        var arrowheadControls = dialogSession.controls.arrowhead;
        var refreshPreview = function () { updatePreview(dialogSession); };
        var i;
        for (i = 0; i < arrowheadControls.favoriteArrowIcons.length; i++) {
            arrowheadControls.favoriteArrowIcons[i].onClick = (function (favoriteIcon, favoriteArrow) {
                return function () {
                    /* ⌘＋option＋クリックは［終点も同じ］（始点と終点のリンク）を切り替える。
                       option＋クリックは始点と終点を入れ替える（終点も同じのときは入れ替えようがないので何もしない）
                       Cmd-Option-click toggles Same at end; Option-click swaps the start and end (nothing to swap while Same at end is on) */
                    var keyboardState = ScriptUI.environment.keyboardState;
                    if (keyboardState.altKey && keyboardState.metaKey) {
                        setLinkToggleValue(arrowheadControls.sameEndToggle, !arrowheadControls.sameEndToggle.value);
                    } else if (keyboardState.altKey && !arrowheadControls.sameEndToggle.value) {
                        setLinkToggleValue(arrowheadControls.swapEndsToggle, !arrowheadControls.swapEndsToggle.value);
                    }
                    selectArrowChoice(dialogSession, favoriteIcon, favoriteArrow);
                };
            })(arrowheadControls.favoriteArrowIcons[i], FAVORITE_ARROWS[i]);
        }
        /* メニューから選んだときも、その他のラジオをオンにする / Picking from the menu turns on its radio */
        arrowheadControls.otherArrowRadio.onClick = arrowheadControls.otherArrowList.onChange = function () {
            selectArrowChoice(dialogSession, arrowheadControls.otherArrowRadio, null);
        };
        /* 入力中はアクションのプレビューを取り消す（確定時に貼り直す）/ drop the action preview while typing; reapplied on commit */
        arrowheadControls.arrowScaleInput.onChanging = function () { dialogSession.previewController.clearAction(); };
        addCommitHandler(arrowheadControls.arrowScaleInput, refreshPreview);
        /* 終点も同じなら両端が同じになるので、入れ替えはディム / Same at both ends makes swapping pointless */
        arrowheadControls.sameEndToggle.onClick = function () {
            syncArrowOptionsEnabled(arrowheadControls);
            refreshPreview();
        };
        arrowheadControls.swapEndsToggle.onClick = refreshPreview;
        for (i = 0; i < arrowheadControls.tipAlignIcons.length; i++) arrowheadControls.tipAlignIcons[i].onClick = refreshPreview;
    }

    /**
     * 破線パネルのイベントをつなぐ。破線・ドット点線を選んだら矢印を［なし］にし、その種類の初期値を入れる
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function bindDashEvents(dialogSession) {
        var dialogControls = dialogSession.controls;
        var dashControls = dialogControls.dash;
        var refreshDashCalc = function () { updateDashCalc(dialogSession); };
        var i;
        dashControls.noDashRadio.onClick = dashControls.dashedRadio.onClick = dashControls.dottedRadio.onClick = function () {
            if (!dashControls.noDashRadio.value) selectNoArrowhead(dialogSession);
            fillDashCalcDefaults(dialogControls, dialogSession.firstPathMetrics);
            syncDashEnabled(dialogControls);
            refreshDashCalc();
        };
        dashControls.gapToDashRadio.onClick = dashControls.dashToGapRadio.onClick = function () {
            syncDashEnabled(dialogControls);
            refreshDashCalc();
        };
        for (i = 0; i < dashControls.adjustDashEndsIcons.length; i++) dashControls.adjustDashEndsIcons[i].onClick = refreshDashCalc;
        var dashInputs = [dashControls.segmentsInput, dashControls.gapInput, dashControls.dashLengthInput];
        for (i = 0; i < dashInputs.length; i++) {
            /* 入力中はアクションのプレビューを取り消す（確定時に貼り直す）/ drop the action preview while typing; reapplied on commit */
            dashInputs[i].onChanging = function () { dialogSession.previewController.clearAction(); };
            addCommitHandler(dashInputs[i], refreshDashCalc);
        }
    }

    /**
     * 矢印を［なし］にする（プレビューは呼び出し側で更新する）
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function selectNoArrowhead(dialogSession) {
        var arrowheadControls = dialogSession.controls.arrowhead;
        var noArrowIndex = findFavoriteArrowIndex(0);
        if (noArrowIndex < 0) return;
        changeArrowStrokeWidthScale(dialogSession, FAVORITE_ARROWS[noArrowIndex].strokeWidthScale || 1);
        selectArrowByNumber(arrowheadControls, 0);
        setNumberFieldValue(arrowheadControls.arrowScaleInput, FAVORITE_ARROWS[noArrowIndex].scale);
        syncArrowOptionsEnabled(arrowheadControls);
    }

    /**
     * プリセットの行のイベントをつなぐ。選んだら読み込む。
     * ドロップダウンの作り直しでも onChange が来るので、書き込み中は読まない
     * @param {Object} dialogSession - showSettingsDialog() のダイアログの状態
     * @returns {void}
     */
    function bindPresetEvents(dialogSession) {
        var dialogControls = dialogSession.controls;
        var presetControls = dialogControls.preset;

        /* ドロップダウンを作り直して名前を選ぶ（onChange は無視させる）/ refill the dropdown, ignoring its onChange */
        function refreshPresetDropdown(selectedName) {
            dialogSession.isApplyingPreset = true;
            fillPresetDropdown(presetControls, selectedName);
            dialogSession.isApplyingPreset = false;
        }

        presetControls.presetDropdown.onChange = function () {
            if (dialogSession.isApplyingPreset) return;
            var presetName = getSelectedPresetName(presetControls);
            setIconButtonEnabled(presetControls.presetDeleteIcon, presetName !== null);
            if (!presetName) return;
            var presetData = loadPresetMap()[presetName];
            if (!presetData) return;
            dialogSession.isApplyingPreset = true;
            applyPresetData(dialogControls, presetData);
            dialogSession.isApplyingPreset = false;
            /* プリセットの線幅は矢印の倍数を掛けたあとの値なので、倍数だけ合わせる / the preset weight already includes the multiplier */
            dialogSession.strokeWidthScale = getArrowStrokeWidthScale(presetData.arrowNumber);
            syncDashEnabled(dialogControls);
            updateDashCalc(dialogSession);
        };
        presetControls.presetSaveIcon.onClick = function () {
            var savedName = saveCurrentPreset(dialogControls);
            if (savedName) refreshPresetDropdown(savedName);
        };
        presetControls.presetDeleteIcon.onClick = function () {
            if (deleteSelectedPreset(presetControls)) refreshPresetDropdown(null);
        };
        refreshPresetDropdown(null);
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /**
     * 線幅・矢印・線端・角の形状を設定する一時アクションのソースを生成する
     * @param {string} setName - アクションセット名（ASCII 推奨）
     * @param {string} actionName - アクション名（ASCII 推奨）
     * @param {Object} strokeSettings - readDialogSettings() の戻り値
     * @returns {string} アクションのソース
     */
    function buildStrokeActionSource(setName, actionName, strokeSettings) {
        return ''
            + '/version 3\n'
            + buildNameLine(setName)
            + '/isOpen 1\n'
            + '/actionCount 1\n'
            + '/action-1 {\n'
            + '\t' + buildNameLine(actionName)
            + '\t/keyIndex 0\n'
            + '\t/colorIndex 0\n'
            + '\t/isOpen 1\n'
            + '\t/eventCount 1\n'
            + '\t/event-1 {\n'
            + '\t\t/useRulersIn1stQuadrant 0\n'
            + '\t\t/internalName (ai_plugin_setStroke)\n'
            + '\t\t/localizedName [ 0 \n\t\t]\n'
            + '\t\t/isOpen 1\n'
            + '\t\t/isOn 1\n'
            + '\t\t/hasDialog 0\n'
            + '\t\t/parameterCount 11\n'
            /* 可変 / variable */
            + buildUnitRealParam(1, KEY_STROKE_WIDTH, strokeSettings.strokeWidth, UNIT_POINT)
            + buildEnumParamByName(2, KEY_CAP, getLabel(strokeSettings.strokeCap.actionName), strokeSettings.strokeCap.value)     /* 線端 / cap */
            + buildEnumParamByName(3, KEY_JOIN, getLabel(strokeSettings.cornerJoin.actionName), strokeSettings.cornerJoin.value)  /* 角の形状 / join */
            + buildUStrParam(6, KEY_ARROW_HEAD_1, strokeSettings.startArrow)
            + buildUStrParam(7, KEY_ARROW_HEAD_2, strokeSettings.endArrow)
            + buildRealParam(8, KEY_ARROW_SCALE_1, strokeSettings.startScale)
            + buildRealParam(9, KEY_ARROW_SCALE_2, strokeSettings.endScale)
            + buildEnumParamByName(10, KEY_TIP_ALIGN, getLabel(strokeSettings.tipAlign.actionName), strokeSettings.tipAlign.value) /* 先端位置 / tip alignment */
            /* 以下は記録した .aia のまま。破線はアクションの後に DOM で設定する / recorded as-is; dashes are set through the DOM afterwards */
            + buildIntParam(4, KEY_DASH_INT, 0)
            + buildBoolParam(5, KEY_DASH_BOOL, 0)
            + buildEnumParam(11, KEY_STROKE_ALIGN, 'e4b8ade5a4ae', 6, 0)                 /* 線の位置: 中央 / center */
            + '\t}\n'
            + '}\n';
    }

    // --- パラメータ組み立てヘルパー / parameter builders ---

    /**
     * パラメータブロックの外枠を作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} body - ブロックの中身
     * @returns {string} パラメータブロック
     */
    function buildParamBlock(index, key, body) {
        return ''
            + '\t\t/parameter-' + index + ' {\n'
            + '\t\t\t/key ' + key + '\n'
            + '\t\t\t/showInPalette 4294967295\n'
            + body
            + '\t\t}\n';
    }

    /**
     * 単位付き実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @param {number} unitCode - 単位コード
     * @returns {string} パラメータブロック
     */
    function buildUnitRealParam(index, key, value, unitCode) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (unit real)\n'
            + '\t\t\t/value ' + toRealString(value) + '\n'
            + '\t\t\t/unit ' + unitCode + '\n');
    }

    /**
     * 列挙パラメータを作る（表示名は16進で渡す）
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} nameHex - 表示名の16進
     * @param {number} byteLength - 表示名のバイト数
     * @param {number} value - 列挙値
     * @returns {string} パラメータブロック
     */
    function buildEnumParam(index, key, nameHex, byteLength, value) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (enumerated)\n'
            + '\t\t\t/name [ ' + byteLength + ' \n\t\t\t\t' + nameHex + '\n\t\t\t]\n'
            + '\t\t\t/value ' + value + '\n');
    }

    /**
     * 列挙パラメータを表示名から作る（16進は自動で作る）
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} displayName - 表示名
     * @param {number} value - 列挙値
     * @returns {string} パラメータブロック
     */
    function buildEnumParamByName(index, key, displayName, value) {
        var nameHex = toActionHex(displayName);
        return buildEnumParam(index, key, nameHex, nameHex.length / 2, value);
    }

    /**
     * 整数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildIntParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (integer)\n\t\t\t/value ' + value + '\n');
    }

    /**
     * 真偽値パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {boolean|number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildBoolParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (boolean)\n\t\t\t/value ' + (value ? 1 : 0) + '\n');
    }

    /**
     * 実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildRealParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (real)\n\t\t\t/value ' + toRealString(value) + '\n');
    }

    /**
     * Unicode 文字列パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} value - 文字列
     * @returns {string} パラメータブロック
     */
    function buildUStrParam(index, key, value) {
        var valueHex = toActionHex(value);
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (ustring)\n'
            + '\t\t\t/value [ ' + (valueHex.length / 2) + ' \n\t\t\t\t' + valueHex + '\n\t\t\t]\n');
    }

    /**
     * アクションセット名・アクション名の /name 行を作る
     * @param {string} displayName - 名前
     * @returns {string} /name 行
     */
    function buildNameLine(displayName) {
        var nameHex = toActionHex(displayName);
        return '/name [ ' + (nameHex.length / 2) + ' \n\t' + nameHex + '\n]\n';
    }

    /**
     * 5 → "5.0" のように、必ず小数点を含む文字列にする
     * @param {number} value - 値
     * @returns {string} 小数点を含む文字列
     */
    function toRealString(value) {
        var realText = String(Number(value));
        if (realText.indexOf('.') === -1 && realText.indexOf('e') === -1) realText += '.0';
        return realText;
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * ［線］の設定アクションを1回実行する。失敗は例外で伝える
     * @param {Object} strokeSettings - readDialogSettings() の戻り値
     * @returns {void}
     */
    function playStrokeAction(strokeSettings) {
        var actionSource = buildStrokeActionSource(ACTION_SET_NAME, ACTION_NAME, strokeSettings);
        if (!runTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME)) {
            throw new Error(getLabel("alert.actionFailed"));
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提条件を確認し、ダイアログの入力を選択中のパスに適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetDocument = app.activeDocument;
        if (targetDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var strokeSettings = showSettingsDialog(targetDocument);
        if (!strokeSettings) return;

        applyStrokeSettings(collectStrokePaths(targetDocument.selection), strokeSettings);
    }

    main();

})();
