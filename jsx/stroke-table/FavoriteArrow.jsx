#target illustrator
#targetengine "FavoriteArrowEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスに、よく使う矢印と線の設定（線幅・線端・角の形状・破線）をまとめて適用します。
矢印はDOMから操作できないため、一時アクションを生成して実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteArrow.md

### Overview

Applies a favorite arrowhead together with the stroke settings — weight, cap, corner and dashes — to the selected paths.
Arrowheads cannot be reached from the DOM, so a temporary action is generated and played instead.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteArrow.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FavoriteArrow";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteArrow.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteArrow.md"; /* README (English) */

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
    var DEFAULT_ARROW_SCALE  = 100;          /* 倍率の初期値（ポップアップメニューの矢印にも使う）/ initial arrowhead scale, also used for the pop-up arrowheads */
    var DEFAULT_STROKE_CAP   = "butt";       /* 線端の初期値（butt / round / projecting）/ default cap */
    var DEFAULT_CORNER_JOIN  = "miter";      /* 角の形状の初期値（miter / round / bevel）/ default join */

    /* ラジオボタンで出す矢印と倍率。ここに無い矢印はポップアップメニューに並ぶ
       Arrowheads offered as radio buttons, with their scales. The rest go in the pop-up menu */
    var FAVORITE_ARROWS = [
        { number: 1,  scale: 100 },
        { number: 8,  scale: 25 },
        { number: 11, scale: 100 },
        { number: 27, scale: 100 }
    ];

    /* 破線のパターン（線幅に対する倍率）。オンにしたときの分割数・間隔・線分の初期値に使う
       丸型線端では、見かけの線分は dash＋線幅、すき間は gap−線幅になる。ドット点線は丸型線端に固定
       Dash patterns (multiples of the stroke weight), used for the initial segments, gap and dash.
       With round caps a visible dash is dash + weight and a visible gap is gap - weight. Dotted always uses round caps */
    var DASH_PATTERNS = {
        dashed: { dash: 2, gap: 3 },         /* 破線 / dashed */
        dotted: { dash: 0, gap: 2 }          /* ドット点線 / dotted */
    };

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
    var STROKE_LABEL_WIDTH = 64;             /* 線パネルの項目名の幅（コロンまで収まる幅）/ stroke panel label width */
    var ARROW_LABEL_WIDTH  = 40;             /* 矢印パネルの項目名の幅 / arrowhead panel label width */
    var DASH_LABEL_WIDTH   = 48;             /* 破線の計算の項目名の幅 / dash calculation label width */
    var FIELD_CHARACTERS   = 4;              /* 数値入力欄の文字数 / numeric field width */

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
            title: { ja: "線と矢印を設定", en: "Set Stroke and Arrowheads" }
        },
        panel: {
            stroke: { ja: "線", en: "Stroke" },
            arrowhead: { ja: "矢印", en: "Arrowheads" },
            dash: { ja: "破線", en: "Dashed Line" },
            dashCalc: { ja: "破線の計算", en: "Dash Calculation" },
            calcMethod: { ja: "計算方法", en: "Calculation" },
            tipAlign: { ja: "先端位置", en: "Tip Alignment" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Weight" },
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
            buttCap: { ja: "なし", en: "Butt" },
            roundCap: { ja: "丸型", en: "Round" },
            projectingCap: { ja: "突出", en: "Projecting" },
            miterJoin: { ja: "マイター", en: "Miter" },
            roundJoin: { ja: "ラウンド", en: "Round" },
            bevelJoin: { ja: "ベベル", en: "Bevel" },
            gapToDash: { ja: "間隔→線分", en: "Gap→Dash" },
            dashToGap: { ja: "線分→間隔", en: "Dash→Gap" },
            tipAtEnd: { ja: "矢印の先端をパスの終点に配置", en: "Place arrow tip at end of path" },
            tipBeyondEnd: { ja: "矢印の先端をパスの終点から配置", en: "Extend arrow tip beyond end of path" }
        },
        checkbox: {
            sameEnd: { ja: "終点も同じ", en: "Same at end" },
            swapEnds: { ja: "始点と終点を入れ替え", en: "Swap start and end" },
            dashed: { ja: "破線", en: "Dashed" },
            dotted: { ja: "ドット点線", en: "Dotted" },
            adjustDashEnds: { ja: "両端を調整", en: "Adjust ends" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            strokeWidth: { ja: "線の太さです。", en: "Weight of the stroke." },
            favoriteArrow: {
                ja: "始点に付ける矢印です。倍率はこの矢印に合わせた値に変わります。",
                en: "Arrowhead for the start of the path. The scale changes to suit it."
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
            dashed: {
                ja: "破線にします。オンにすると、線幅から決めた分割数・間隔・線分が入ります。",
                en: "Makes a dashed line. Turning it on fills in Segments, Gap and Dash based on the stroke weight."
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
            tipAlign: { ja: "矢印をパスの端にどう合わせるかです。", en: "How the arrowhead lines up with the end of the path." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
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
        /* アクションに埋め込む Illustrator の表示名。矢印名はラジオボタン・メニューの表示にも使う
           Illustrator labels embedded in the action; arrowhead names double as radio and menu labels */
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
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

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

    /* 先端位置・線端・角の形状の選択肢。value はアクションの enumerated 値
       線端・角の形状は表示名がそのまま説明になるので、helpTip にも actionName を使う
       Tip alignment, cap and join options; value is the action's enumerated value.
       The cap and join display names explain themselves, so they double as helpTips */
    var TIP_ALIGN_OPTIONS = [
        { key: "atEnd",     label: "radio.tipAtEnd",     actionName: "actionName.tipAtEnd",     value: 0 },
        { key: "beyondEnd", label: "radio.tipBeyondEnd", actionName: "actionName.tipBeyondEnd", value: 1 }
    ];
    var STROKE_CAP_OPTIONS = [
        { key: "butt",       label: "radio.buttCap",       actionName: "actionName.buttCap",       value: 0 },
        { key: "round",      label: "radio.roundCap",      actionName: "actionName.roundCap",      value: 1 },
        { key: "projecting", label: "radio.projectingCap", actionName: "actionName.projectingCap", value: 2 }
    ];
    var CORNER_JOIN_OPTIONS = [
        { key: "miter", label: "radio.miterJoin", actionName: "actionName.miterJoin", value: 0 },
        { key: "round", label: "radio.roundJoin", actionName: "actionName.roundJoin", value: 1 },
        { key: "bevel", label: "radio.bevelJoin", actionName: "actionName.bevelJoin", value: 2 }
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
     * ラジオボタンに出す矢印名の一覧を作る（FAVORITE_ARROWS の順）
     * @returns {string[]} 矢印名
     */
    function buildFavoriteArrowNames() {
        var favoriteArrowNames = [];
        for (var i = 0; i < FAVORITE_ARROWS.length; i++) {
            favoriteArrowNames.push(getLabel("actionName.arrowheadPrefix") + FAVORITE_ARROWS[i].number);
        }
        return favoriteArrowNames;
    }

    /**
     * ラジオボタンに出さない矢印名の一覧を作る（番号順）
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
     * 破線を DOM で設定する。アクションが破線を解除するので、アクションの後に呼ぶ
     * @param {PathItem[]} strokePaths - collectStrokePaths() で集めたパス
     * @param {Object|null} dashCalc - readDashCalc() の戻り値。null なら何もしない
     * @returns {void}
     */
    function applyDashStyle(strokePaths, dashCalc) {
        if (!dashCalc) return;
        for (var i = 0; i < strokePaths.length; i++) {
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
        playStrokeAction(strokeSettings);
        applyDashStyle(strokePaths, strokeSettings.dashCalc);
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
                if (originalState.stroked) strokePaths[i].strokeWidth = originalState.strokeWidth;
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
     * 右揃えの項目名を持つ行を追加する
     * @param {Panel|Group} parent - 追加先
     * @param {string} labelKey - 項目名の LABELS パス
     * @param {number} labelWidth - 項目名の幅
     * @returns {Group} 追加した行
     */
    function addLabeledRow(parent, labelKey, labelWidth) {
        var labeledRow = parent.add("group");
        setupRow(labeledRow, "left", ROW_SPACING);
        var rowLabel = labeledRow.add("statictext", undefined, labelText(labelKey));
        rowLabel.preferredSize.width = labelWidth;
        rowLabel.justify = "right";
        return labeledRow;
    }

    /**
     * 幅いっぱいに広げないチェックボックスを追加する
     * @param {Panel|Group} parent - 追加先
     * @param {string} labelKey - 表示名の LABELS パス
     * @param {string} tooltipKey - helpTip の LABELS パス
     * @param {boolean} isChecked - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parent, labelKey, tooltipKey, isChecked) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel(labelKey));
        optionCheckbox.helpTip = getLabel(tooltipKey);
        optionCheckbox.value = isChecked;
        optionCheckbox.alignment = "left";
        return optionCheckbox;
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
     * 選択肢の一覧からラジオボタンを並べる。各ボタンの optionDefinition に選択肢を持たせる
     * @param {Panel|Group} parent - 追加先
     * @param {Object[]} optionDefinitions - 選択肢（key / label / actionName）
     * @param {string} selectedKey - 初期選択の key
     * @param {string} [tooltipKey] - helpTip の LABELS パス（省略時は選択肢の actionName）
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addOptionRadios(parent, optionDefinitions, selectedKey, tooltipKey) {
        var optionRadios = [];
        for (var i = 0; i < optionDefinitions.length; i++) {
            var optionDefinition = optionDefinitions[i];
            var optionRadio = addOptionRadio(parent, getLabel(optionDefinition.label),
                getLabel(tooltipKey || optionDefinition.actionName));
            optionRadio.optionDefinition = optionDefinition;
            optionRadio.value = (optionDefinition.key === selectedKey);
            optionRadios.push(optionRadio);
        }
        return optionRadios;
    }

    /**
     * オンになっているラジオボタンの選択肢を返す
     * @param {RadioButton[]} optionRadios - addOptionRadios() で作ったラジオボタン
     * @returns {Object} 選択肢。どれもオフなら先頭
     */
    function getCheckedOption(optionRadios) {
        for (var i = 0; i < optionRadios.length; i++) {
            if (optionRadios[i].value) return optionRadios[i].optionDefinition;
        }
        return optionRadios[0].optionDefinition;
    }

    /**
     * 項目名・∧∨・入力欄・単位をひと組にした数値欄を追加する（ステップボタン部品の addSteppedField を使う）
     * @param {Panel|Group} parent - 追加先
     * @param {Object} fieldOptions - labelKey / labelWidth / text / min / integer / unitKey / tooltipKey
     * @returns {EditText} 入力欄
     */
    function addNumberField(parent, fieldOptions) {
        var numberInput = addSteppedField(parent, {
            label: labelText(fieldOptions.labelKey),
            labelWidth: fieldOptions.labelWidth,
            text: fieldOptions.text,
            characters: FIELD_CHARACTERS,
            min: fieldOptions.min,
            integer: fieldOptions.integer,
            onStep: notifySteppedInput
        });
        numberInput.helpTip = getLabel(fieldOptions.tooltipKey);
        /* 行の間隔をほかの行にそろえ、単位を後ろに足す / match the other rows and append the unit */
        var fieldRow = numberInput.parent.parent;
        fieldRow.spacing = ROW_SPACING;
        if (fieldOptions.unitKey) fieldRow.add("statictext", undefined, getLabel(fieldOptions.unitKey));
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
            labelKey: "fieldLabel.strokeWidth", labelWidth: STROKE_LABEL_WIDTH, text: String(getDefaultStrokeWidth()),
            min: 0, unitKey: "unit.point", tooltipKey: "tooltip.strokeWidth"
        });
        var strokeCapRow = addLabeledRow(strokePanel, "fieldLabel.strokeCap", STROKE_LABEL_WIDTH);
        var strokeCapRadios = addOptionRadios(strokeCapRow, STROKE_CAP_OPTIONS, DEFAULT_STROKE_CAP);
        var cornerJoinRow = addLabeledRow(strokePanel, "fieldLabel.cornerJoin", STROKE_LABEL_WIDTH);
        var cornerJoinRadios = addOptionRadios(cornerJoinRow, CORNER_JOIN_OPTIONS, DEFAULT_CORNER_JOIN);

        return {
            strokeWidthInput: strokeWidthInput,
            strokeCapRow: strokeCapRow,
            strokeCapRadios: strokeCapRadios,
            cornerJoinRadios: cornerJoinRadios
        };
    }

    /**
     * 矢印パネル（よく使う矢印・その他の矢印・倍率・終点の扱い）を作る
     * @param {Window|Group} parent - 追加先
     * @returns {Object} 矢印パネルのコントロール
     */
    function buildArrowheadPanel(parent) {
        var arrowheadPanel = parent.add("panel", undefined, getLabel("panel.arrowhead"));
        setupPanel(arrowheadPanel, 6);

        var favoriteArrowNames = buildFavoriteArrowNames();
        var favoriteArrowRadios = [];
        for (var i = 0; i < favoriteArrowNames.length; i++) {
            favoriteArrowRadios.push(addOptionRadio(arrowheadPanel, favoriteArrowNames[i], getLabel("tooltip.favoriteArrow")));
        }
        favoriteArrowRadios[0].value = true;

        /* その他：ラジオ＋ポップアップメニュー。別グループのラジオは排他にならないので、onClick で切り替える
           Others: radio + pop-up. Radios in another group are not exclusive, so onClick handles it */
        var otherArrowRow = arrowheadPanel.add("group");
        setupRow(otherArrowRow, "left", 4);
        var otherArrowTooltip = getLabel("tooltip.otherArrow", { scale: DEFAULT_ARROW_SCALE });
        var otherArrowRadio = otherArrowRow.add("radiobutton", undefined, "");
        otherArrowRadio.helpTip = otherArrowTooltip;
        var otherArrowList = otherArrowRow.add("dropdownlist", undefined, buildOtherArrowNames());
        otherArrowList.helpTip = otherArrowTooltip;
        otherArrowList.selection = 0;

        /* 0 以下は既定の倍率に戻されるので、下限は 1 / values of 0 or less fall back to the default, so the minimum is 1 */
        var arrowScaleInput = addNumberField(arrowheadPanel, {
            labelKey: "fieldLabel.arrowScale", labelWidth: ARROW_LABEL_WIDTH, text: String(FAVORITE_ARROWS[0].scale),
            min: 1, unitKey: "unit.percent", tooltipKey: "tooltip.arrowScale"
        });

        return {
            favoriteArrowRadios: favoriteArrowRadios,
            otherArrowRadio: otherArrowRadio,
            otherArrowList: otherArrowList,
            arrowScaleInput: arrowScaleInput,
            sameEndCheckbox: addOptionCheckbox(arrowheadPanel, "checkbox.sameEnd", "tooltip.sameEnd", false),
            swapEndsCheckbox: addOptionCheckbox(arrowheadPanel, "checkbox.swapEnds", "tooltip.swapEnds", false)
        };
    }

    /**
     * 破線パネル（破線の種類・破線の計算・計算方法・両端を調整）を作る
     * @param {Window|Group} parent - 追加先
     * @returns {Object} 破線パネルのコントロール
     */
    function buildDashPanel(parent) {
        var dashPanel = parent.add("panel", undefined, getLabel("panel.dash"));
        setupPanel(dashPanel, 6);

        var dashedCheckbox = addOptionCheckbox(dashPanel, "checkbox.dashed", "tooltip.dashed", false);
        var dottedCheckbox = addOptionCheckbox(dashPanel, "checkbox.dotted", "tooltip.dotted", false);

        /* 破線の計算（DashGapCalculator から移植）/ Dash calculation (ported from DashGapCalculator) */
        var dashCalcPanel = dashPanel.add("panel", undefined, getLabel("panel.dashCalc"));
        setupPanel(dashCalcPanel, 6);
        var segmentsInput = addNumberField(dashCalcPanel, {
            labelKey: "fieldLabel.segments", labelWidth: DASH_LABEL_WIDTH, text: "1",
            min: 1, integer: true, tooltipKey: "tooltip.segments"
        });
        var gapInput = addNumberField(dashCalcPanel, {
            labelKey: "fieldLabel.gap", labelWidth: DASH_LABEL_WIDTH, text: "0",
            min: 0, unitKey: "unit.point", tooltipKey: "tooltip.gap"
        });
        var dashLengthInput = addNumberField(dashCalcPanel, {
            labelKey: "fieldLabel.dash", labelWidth: DASH_LABEL_WIDTH, text: "0",
            min: 0, unitKey: "unit.point", tooltipKey: "tooltip.dash"
        });

        var calcMethodPanel = dashPanel.add("panel", undefined, getLabel("panel.calcMethod"));
        setupPanel(calcMethodPanel, 6);
        var gapToDashRadio = addOptionRadio(calcMethodPanel, getLabel("radio.gapToDash"), getLabel("tooltip.gapToDash"));
        var dashToGapRadio = addOptionRadio(calcMethodPanel, getLabel("radio.dashToGap"), getLabel("tooltip.dashToGap"));
        gapToDashRadio.value = true;

        return {
            dashedCheckbox: dashedCheckbox,
            dottedCheckbox: dottedCheckbox,
            dashCalcPanel: dashCalcPanel,
            segmentsInput: segmentsInput,
            gapInput: gapInput,
            dashLengthInput: dashLengthInput,
            calcMethodPanel: calcMethodPanel,
            gapToDashRadio: gapToDashRadio,
            dashToGapRadio: dashToGapRadio,
            adjustDashEndsCheckbox: addOptionCheckbox(dashPanel, "checkbox.adjustDashEnds", "tooltip.adjustDashEnds", true)
        };
    }

    /**
     * 設定用ダイアログのコントロールをすべて作る（イベントは設定しない）
     * @returns {Object} ダイアログとパネルごとのコントロール
     */
    function buildSettingsDialog() {
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        var strokeControls = buildStrokePanel(settingsDialog);

        /* 矢印と破線を 2 カラムで並べる / Lay out arrowheads and dashes in two columns */
        var arrowDashColumns = settingsDialog.add("group");
        arrowDashColumns.orientation = "row";
        arrowDashColumns.alignChildren = ["fill", "fill"];
        arrowDashColumns.spacing = COLUMN_SPACING;
        var arrowheadControls = buildArrowheadPanel(arrowDashColumns);
        var dashControls = buildDashPanel(arrowDashColumns);

        var tipAlignPanel = settingsDialog.add("panel", undefined, getLabel("panel.tipAlign"));
        setupPanel(tipAlignPanel, 6);
        var tipAlignRadios = addOptionRadios(tipAlignPanel, TIP_ALIGN_OPTIONS, TIP_ALIGN_OPTIONS[0].key, "tooltip.tipAlign");

        /* ボタン（左：プレビュー／中央：スペーサー／右：キャンセル・OK） */
        /* Buttons (left: preview, center: spacer, right: cancel and OK) */
        var buttonRow = addButtonRow(settingsDialog);
        var previewCheckbox = addOptionCheckbox(buttonRow.leftGroup, "checkbox.preview", "tooltip.preview", false);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            settingsDialog: settingsDialog,
            stroke: strokeControls,
            arrowhead: arrowheadControls,
            dash: dashControls,
            tipAlignRadios: tipAlignRadios,
            previewCheckbox: previewCheckbox
        };
    }

    // =========================================
    // ダイアログの入力値 / Dialog values
    // =========================================

    /**
     * 選択中の矢印名を返す（ラジオボタンとメニューの表示名がそのまま矢印名）
     * @param {Object} arrowheadControls - buildArrowheadPanel() の戻り値
     * @returns {string} 矢印名
     */
    function getSelectedArrowName(arrowheadControls) {
        if (arrowheadControls.otherArrowRadio.value) return arrowheadControls.otherArrowList.selection.text;
        var favoriteArrowRadios = arrowheadControls.favoriteArrowRadios;
        for (var i = 0; i < favoriteArrowRadios.length; i++) {
            if (favoriteArrowRadios[i].value) return favoriteArrowRadios[i].text;
        }
        return favoriteArrowRadios[0].text;
    }

    /**
     * 破線の計算の入力値を読む
     * @param {Object} dashControls - buildDashPanel() の戻り値
     * @returns {Object|null} calcDashArray() に渡す値。破線がオフなら null
     */
    function readDashCalc(dashControls) {
        if (!dashControls.dashedCheckbox.value && !dashControls.dottedCheckbox.value) return null;
        return {
            style: dashControls.dottedCheckbox.value ? "dotted" : "dashed",
            mode: dashControls.dashToGapRadio.value ? "dashToGap" : "gapToDash",
            segments: parseInt(dashControls.segmentsInput.text, 10),
            gapPt: parseFloat(dashControls.gapInput.text),
            dashPt: parseFloat(dashControls.dashLengthInput.text),
            adjustEnds: dashControls.adjustDashEndsCheckbox.value
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
        var isSwapped = !arrowheadControls.sameEndCheckbox.value && arrowheadControls.swapEndsCheckbox.value;
        var hasEndArrow = arrowheadControls.sameEndCheckbox.value || isSwapped;

        var dashCalc = readDashCalc(dialogControls.dash);
        /* ドット点線は長さ0の線分なので、丸型線端でないと見えない / Dots are zero-length dashes, visible only with round caps */
        var isDotted = (dashCalc !== null && dashCalc.style === "dotted");

        return {
            strokeWidth: strokeWidth,
            startArrow: isSwapped ? noArrowName : arrowName,
            startScale: isSwapped ? DEFAULT_ARROW_SCALE : arrowScale,
            endArrow: hasEndArrow ? arrowName : noArrowName,
            endScale: hasEndArrow ? arrowScale : DEFAULT_ARROW_SCALE,
            tipAlign: getCheckedOption(dialogControls.tipAlignRadios),
            strokeCap: isDotted ? findOptionByKey(STROKE_CAP_OPTIONS, "round") : getCheckedOption(dialogControls.stroke.strokeCapRadios),
            cornerJoin: getCheckedOption(dialogControls.stroke.cornerJoinRadios),
            dashCalc: dashCalc
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
     * 破線の有無・種類・計算方法に合わせて有効／無効を切り替える。
     * 入力する側の欄だけ有効にし、ドット点線は分割数だけ。ドット点線の間は線端を丸型に固定するのでディム
     * @param {Object} dialogControls - buildSettingsDialog() の戻り値
     * @returns {void}
     */
    function syncDashEnabled(dialogControls) {
        var dashControls = dialogControls.dash;
        var isDotted = dashControls.dottedCheckbox.value;
        var hasDash = dashControls.dashedCheckbox.value || isDotted;
        var isDashed = hasDash && !isDotted;
        dialogControls.stroke.strokeCapRow.enabled = !isDotted;
        dashControls.dashCalcPanel.enabled = hasDash;
        dashControls.calcMethodPanel.enabled = isDashed;
        dashControls.adjustDashEndsCheckbox.enabled = hasDash;
        setSteppedFieldEnabled(dashControls.segmentsInput, hasDash);
        setSteppedFieldEnabled(dashControls.gapInput, isDashed && dashControls.gapToDashRadio.value);
        setSteppedFieldEnabled(dashControls.dashLengthInput, isDashed && dashControls.dashToGapRadio.value);
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
        var firstPathMetrics = getFirstPathMetrics(collectStrokePaths(targetDocument.selection));
        var previewController = createPreviewController();
        var dialogControls = buildSettingsDialog();
        var strokeControls = dialogControls.stroke;
        var arrowheadControls = dialogControls.arrowhead;
        var dashControls = dialogControls.dash;
        var previewCheckbox = dialogControls.previewCheckbox;
        var i;

        /* プレビューを現在の入力値で更新する（矢印を含むためアクションを実行） */
        /* Refresh the preview with the current values (plays the action for arrowheads) */
        function updatePreview() {
            if (!previewCheckbox.value) {
                previewController.reset();
                return;
            }
            var strokeSettings = readDialogSettings(dialogControls);
            if (strokeSettings) previewController.previewSettings(strokeSettings);
        }

        /* 入力中はアクションのプレビューを取り消す（確定時に貼り直す） */
        /* Drop the action preview while editing; it is reapplied on commit */
        function invalidatePreview() {
            if (previewCheckbox.value) previewController.clearAction();
        }

        /* 破線の計算を更新してプレビューする / Recompute the dash fields and preview */
        function updateDashCalc() {
            refreshDashCalcResult(dashControls, firstPathMetrics);
            updatePreview();
        }

        /* 破線とドット点線は片方だけ。オンにしたらもう一方を外し、初期値を入れる
           Only one dash style; turning one on clears the other and fills in the values */
        function toggleDashStyle(clickedCheckbox, otherCheckbox) {
            if (clickedCheckbox.value) otherCheckbox.value = false;
            fillDashCalcDefaults(dialogControls, firstPathMetrics);
            syncDashEnabled(dialogControls);
            updateDashCalc();
        }

        /* 矢印のラジオを1つだけ選び、倍率を入れる / Select one arrowhead radio and set its scale */
        function selectArrowRadio(targetRadio, arrowScale) {
            var favoriteArrowRadios = arrowheadControls.favoriteArrowRadios;
            for (var j = 0; j < favoriteArrowRadios.length; j++) favoriteArrowRadios[j].value = (favoriteArrowRadios[j] === targetRadio);
            arrowheadControls.otherArrowRadio.value = (arrowheadControls.otherArrowRadio === targetRadio);
            setNumberFieldValue(arrowheadControls.arrowScaleInput, arrowScale);
            updatePreview();
        }

        /* 線 / Stroke：線幅は入力中も DOM で即時プレビュー、確定時にアクションで貼り直す */
        strokeControls.strokeWidthInput.onChanging = function () {
            if (previewCheckbox.value) previewController.previewStrokeWidth(parseFloat(strokeControls.strokeWidthInput.text));
        };
        addCommitHandler(strokeControls.strokeWidthInput, updatePreview);
        var optionRadios = strokeControls.strokeCapRadios.concat(strokeControls.cornerJoinRadios, dialogControls.tipAlignRadios);
        for (i = 0; i < optionRadios.length; i++) optionRadios[i].onClick = updatePreview;

        /* 矢印 / Arrowheads：形状を選んだら、その矢印の倍率を入れる */
        for (i = 0; i < arrowheadControls.favoriteArrowRadios.length; i++) {
            arrowheadControls.favoriteArrowRadios[i].onClick = (function (favoriteRadio, favoriteArrow) {
                return function () { selectArrowRadio(favoriteRadio, favoriteArrow.scale); };
            })(arrowheadControls.favoriteArrowRadios[i], FAVORITE_ARROWS[i]);
        }
        /* メニューから選んだときも、その他のラジオをオンにする / Picking from the menu turns on its radio */
        arrowheadControls.otherArrowRadio.onClick = arrowheadControls.otherArrowList.onChange = function () {
            selectArrowRadio(arrowheadControls.otherArrowRadio, DEFAULT_ARROW_SCALE);
        };
        arrowheadControls.arrowScaleInput.onChanging = invalidatePreview;
        addCommitHandler(arrowheadControls.arrowScaleInput, updatePreview);
        /* 終点も同じなら両端が同じになるので、入れ替えはディム / Same at both ends makes swapping pointless */
        arrowheadControls.sameEndCheckbox.onClick = function () {
            arrowheadControls.swapEndsCheckbox.enabled = !arrowheadControls.sameEndCheckbox.value;
            updatePreview();
        };
        arrowheadControls.swapEndsCheckbox.onClick = updatePreview;

        /* 破線 / Dashes */
        dashControls.dashedCheckbox.onClick = function () { toggleDashStyle(dashControls.dashedCheckbox, dashControls.dottedCheckbox); };
        dashControls.dottedCheckbox.onClick = function () { toggleDashStyle(dashControls.dottedCheckbox, dashControls.dashedCheckbox); };
        dashControls.gapToDashRadio.onClick = dashControls.dashToGapRadio.onClick = function () {
            syncDashEnabled(dialogControls);
            updateDashCalc();
        };
        dashControls.adjustDashEndsCheckbox.onClick = updateDashCalc;
        var dashInputs = [dashControls.segmentsInput, dashControls.gapInput, dashControls.dashLengthInput];
        for (i = 0; i < dashInputs.length; i++) {
            dashInputs[i].onChanging = invalidatePreview;
            addCommitHandler(dashInputs[i], updateDashCalc);
        }
        syncDashEnabled(dialogControls);

        previewCheckbox.onClick = updatePreview;

        prepareDialogWindow(dialogControls.settingsDialog, SCRIPT_NAME);
        var isAccepted = (dialogControls.settingsDialog.show() === 1);

        /* プレビューを必ず取り消してから本適用に進む / Always revert the preview before applying */
        previewController.reset();
        if (!isAccepted) return null;

        var strokeSettings = readDialogSettings(dialogControls);
        if (!strokeSettings) alert(getLabel("alert.invalidWidth"));
        return strokeSettings;
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
